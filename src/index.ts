import "dotenv/config";
import "isomorphic-fetch";

import fs from "fs";
import https from "https";
import path from "path";
import express from "express";

import { App, ExpressAdapter, IPlugin } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : undefined,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : undefined,
};

const adapter = new ExpressAdapter();

if (sslOptions.cert && sslOptions.key) {
  const httpsServer = https.createServer(sslOptions, (adapter as any).express);
  (adapter as any).server = httpsServer;
}

const plugins: IPlugin[] = [];

if (process.env.SSL_KEY_FILE) {
  plugins.push(new DevtoolsPlugin());
}

const app = new App({
  logger: new ConsoleLogger("tab", { level: "debug" }),
  plugins,
  httpServerAdapter: adapter,
});

const expressApp = (adapter as any).express;

expressApp.use(express.json());

function graphIsConfigured() {
  const values = [
    process.env.AAD_APP_TENANT_ID,
    process.env.AAD_APP_CLIENT_ID,
    process.env.AAD_APP_CLIENT_SECRET,
  ];

  return values.every(
    (value) => value && !value.includes("REPLACE_WITH")
  );
}

async function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_APP_TENANT_ID!,
    process.env.AAD_APP_CLIENT_ID!,
    process.env.AAD_APP_CLIENT_SECRET!
  );

  const token = await credential.getToken("https://graph.microsoft.com/.default");

  return Client.init({
    authProvider: (done) => {
      done(null, token?.token || null);
    },
  });
}

const ALLOWED_EMAILS = new Set([
  "borja.giraldez@dehn.de",
  "alberto.reino@dehn.de",
  "rasciel.villegas@dehn.de",
  "zigor.lopez@dehn.de",
  "jose.ruano@dehn.de",
]);

async function loadChannelMembers(
  graphClient: Client,
  teamId: string,
  channelId: string
): Promise<ChannelMember[]> {
  const result = await graphClient
    .api(`/teams/${teamId}/channels/${channelId}/members`)
    .get();

  console.log("Raw channel members:", JSON.stringify(result.value, null, 2));

  const members: ChannelMember[] = [];

  for (const member of result.value) {
    const userId = member.userId;

    let email: string | undefined;
    let displayName = member.displayName || "Unnamed user";

    if (userId) {
      try {
        const user = await graphClient
          .api(`/users/${userId}`)
          .select("id,displayName,mail,userPrincipalName")
          .get();

        displayName = user.displayName || displayName;
        email = user.mail || user.userPrincipalName;
      } catch (error) {
        console.warn(`Could not load user details for ${displayName}`, error);
      }
    }

    members.push({
      id: member.id,
      userId,
      displayName,
      email,
    });
  }

  return members.filter(
    (m) => m.email && ALLOWED_EMAILS.has(m.email.toLowerCase())
  );
}

async function areParticipantsFree(
  graphClient: Client,
  participants: ChannelMember[],
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
  durationMinutes: number
): Promise<boolean> {
  const organizer = participants[0];

  // availabilityViewInterval must be between 5 and 1440 AND strictly less than the window length
  const availabilityViewInterval = Math.max(5, Math.floor(durationMinutes / 2));

  const result = await graphClient
    .api(`/users/${organizer.userId}/calendar/getSchedule`)
    .post({
      schedules: participants.map((p) => p.email),
      startTime: {
        dateTime: startDateTime,
        timeZone,
      },
      endTime: {
        dateTime: endDateTime,
        timeZone,
      },
      availabilityViewInterval,
    });

  return result.value.every((schedule: any) => {
    return !schedule.scheduleItems || schedule.scheduleItems.length === 0;
  });
}

async function findAvailableSlot(
  graphClient: Client,
  participants: ChannelMember[],
  durationMinutes: number,
  timeZone: string
): Promise<{ startDateTime: string; endDateTime: string } | null> {
  const { start, end } = buildSearchWindowFromNow();

  let cursor = new Date(start);

  while (cursor < end) {
    const { weekday, hour } = getZonedParts(cursor, timeZone);

    const isWeekday = weekday !== 0 && weekday !== 6;
    const isWorkingHour = hour >= 9 && hour < 17;

    if (isWeekday && isWorkingHour) {
      const slotEnd = addMinutes(cursor, durationMinutes);

      const startDateTime = toGraphDateTime(cursor, timeZone);
      const endDateTime = toGraphDateTime(slotEnd, timeZone);

      const free = await areParticipantsFree(
        graphClient,
        participants,
        startDateTime,
        endDateTime,
        timeZone,
        durationMinutes
      );

      if (free) {
        return { startDateTime, endDateTime };
      }
    }

    cursor = addMinutes(cursor, 30);
  }

  return null;
}

async function createCalendarEvent(
  graphClient: Client,
  participants: ChannelMember[],
  subject: string,
  startDateTime: string,
  endDateTime: string,
  timeZone: string
) {
  const organizer = participants[0];

  const attendees = participants.slice(1).map((p) => ({
    emailAddress: {
      address: p.email,
      name: p.displayName,
    },
    type: "required",
  }));

  return graphClient
    .api(`/users/${organizer.userId}/events`)
    .post({
      subject,
      body: {
        contentType: "HTML",
        content: "Created by DDS Coffee Talk.",
      },
      start: {
        dateTime: startDateTime,
        timeZone,
      },
      end: {
        dateTime: endDateTime,
        timeZone,
      },
      attendees,
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    });
}

expressApp.get("/api/status", (_req: any, res: any) => {
  res.json({
    ok: true,
    message: "Backend is connected",
    graphConfigured: graphIsConfigured(),
  });
});

expressApp.post("/api/channel-members", async (req: any, res: any) => {
  try {
    const { teamId, channelId } = req.body;

    if (!teamId || !channelId) {
      return res.status(400).json({
        ok: false,
        message: "Open the app as a channel tab to get team and channel context.",
        members: [],
      });
    }

    if (!graphIsConfigured()) {
      return res.json({
        ok: false,
        graphConfigured: false,
        message: "Still pending to configure Graph access.",
        members: [],
      });
    }

    const graphClient = await getGraphClient();
    const members = await loadChannelMembers(graphClient, teamId, channelId);

    return res.json({
      ok: true,
      graphConfigured: true,
      message: `Loaded ${members.length} channel members.`,
      members,
    });
  } catch (error: any) {
    console.error("Graph error:", error);

    return res.status(500).json({
      ok: false,
      graphConfigured: true,
      message: "Graph call failed. Check permissions, tenant, secret, and admin consent.",
      error: error.message,
      members: [],
    });
  }
});

expressApp.post("/api/random-meetings/now", async (req: any, res: any) => {
  try {
    const {
      teamId,
      channelId,
      meetingCount,
      participantsPerMeeting,
      durationMinutes,
      timeZone,
    } = req.body;

    if (!teamId || !channelId) {
      return res.status(400).json({
        ok: false,
        message: "Open the app as a channel tab.",
        meetings: [],
      });
    }

    if (!graphIsConfigured()) {
      return res.json({
        ok: false,
        graphConfigured: false,
        message: "Still pending to configure Graph access.",
        meetings: [],
      });
    }

    const graphClient = await getGraphClient();
    const members = await loadChannelMembers(graphClient, teamId, channelId);

    const groups = createRandomGroups(
      members,
      Number(meetingCount),
      Number(participantsPerMeeting)
    );

    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];
      const slot = await findAvailableSlot(
        graphClient,
        participants,
        Number(durationMinutes),
        timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid"
      );

      if (!slot) {
        meetings.push({
          subject: `DDS Coffee Talk #${i + 1}`,
          participants,
          status: "failed",
          message: "No available slot found.",
        });
        continue;
      }

      const event = await createCalendarEvent(
        graphClient,
        participants,
        `DDS Coffee Talk #${i + 1}`,
        slot.startDateTime,
        slot.endDateTime,
        timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid"
      );

      meetings.push({
        subject: event.subject,
        participants,
        startDateTime: slot.startDateTime,
        endDateTime: slot.endDateTime,
        status: "created",
        message: "Meeting created.",
        webLink: event.webLink,
      });
    }

    return res.json({
      ok: true,
      message: `Created ${meetings.filter((m) => m.status === "created").length} meetings.`,
      meetings,
    });
  } catch (error: any) {
    console.error("Scheduling error:", error);

    return res.status(500).json({
      ok: false,
      message: "Scheduling failed. Check Graph permissions and backend logs.",
      error: error.message,
      meetings: [],
    });
  }
});

expressApp.post("/api/random-meetings/at-time", async (req: any, res: any) => {
  try {
    const {
      teamId,
      channelId,
      meetingCount,
      participantsPerMeeting,
      durationMinutes,
      startDateTime,
      timeZone,
    } = req.body;

    if (!teamId || !channelId || !startDateTime) {
      return res.status(400).json({
        ok: false,
        message: "Missing team/channel context or start date/time.",
        meetings: [],
      });
    }

    if (!graphIsConfigured()) {
      return res.json({
        ok: false,
        graphConfigured: false,
        message: "Still pending to configure Graph access.",
        meetings: [],
      });
    }

    const tz = timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid";
    const graphClient = await getGraphClient();
    const members = await loadChannelMembers(graphClient, teamId, channelId);

    const groups = createRandomGroups(
      members,
      Number(meetingCount),
      Number(participantsPerMeeting)
    );

    // The picked value (e.g. from <input type="datetime-local">) is a naive
    // wall clock the user means in `tz` — interpret it there, not server-local.
    const start = zonedWallClockToInstant(startDateTime, tz);
    const end = addMinutes(start, Number(durationMinutes));

    const startGraph = toGraphDateTime(start, tz);
    const endGraph = toGraphDateTime(end, tz);

    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];

      const free = await areParticipantsFree(
        graphClient,
        participants,
        startGraph,
        endGraph,
        tz,
        Number(durationMinutes)
      );

      if (!free) {
        meetings.push({
          subject: `DDS Coffee Talk #${i + 1}`,
          participants,
          startDateTime: startGraph,
          endDateTime: endGraph,
          status: "failed",
          message: "One or more participants are not available.",
        });
        continue;
      }

      const event = await createCalendarEvent(
        graphClient,
        participants,
        `DDS Coffee Talk #${i + 1}`,
        startGraph,
        endGraph,
        tz
      );

      meetings.push({
        subject: event.subject,
        participants,
        startDateTime: startGraph,
        endDateTime: endGraph,
        status: "created",
        message: "Meeting created.",
        webLink: event.webLink,
      });
    }

    return res.json({
      ok: true,
      message: `Created ${meetings.filter((m) => m.status === "created").length} meetings.`,
      meetings,
    });
  } catch (error: any) {
    console.error("Scheduling error:", error);

    return res.status(500).json({
      ok: false,
      message: "Scheduling failed. Check Graph permissions and backend logs.",
      error: error.message,
      meetings: [],
    });
  }
});

type ChannelMember = {
  id?: string;
  userId?: string;
  displayName: string;
  email?: string;
};

type PlannedMeeting = {
  subject: string;
  participants: ChannelMember[];
  startDateTime?: string;
  endDateTime?: string;
  status: "planned" | "created" | "pending_graph" | "failed";
  message: string;
  webLink?: string;
};

function shuffle<T>(items: T[]): T[] {
  return [...items].sort(() => Math.random() - 0.5);
}

function createRandomGroups(
  members: ChannelMember[],
  meetingCount: number,
  participantsPerMeeting: number
): ChannelMember[][] {
  const usableMembers = members.filter((m) => m.email);
  const shuffled = shuffle(usableMembers);

  const groups: ChannelMember[][] = [];
  let index = 0;

  for (let i = 0; i < meetingCount; i++) {
    const group = shuffled.slice(index, index + participantsPerMeeting);

    if (group.length < participantsPerMeeting) {
      break;
    }

    groups.push(group);
    index += participantsPerMeeting;
  }

  return groups;
}

function addMinutes(date: Date, minutes: number): Date {
  return new Date(date.getTime() + minutes * 60_000);
}

// --- Time zone helpers ---------------------------------------------------
// All scheduling reasoning happens in one explicit IANA time zone (the user's),
// never the server process's ambient zone. JS Date.getHours()/getDay() read the
// process-local zone while toISOString() reads UTC; mixing those two with a
// third zone label (sent to Graph) is what created meetings outside working
// hours. These helpers keep the target zone as the single source of truth.

type ZonedParts = {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  weekday: number; // 0 = Sunday ... 6 = Saturday
};

function getZonedParts(date: Date, timeZone: string): ZonedParts {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
    weekday: "short",
  }).formatToParts(date);

  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? "";

  const weekdayMap: Record<string, number> = {
    Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6,
  };

  return {
    year: Number(get("year")),
    month: Number(get("month")),
    day: Number(get("day")),
    // "24" can appear at midnight with hour12:false in some engines; normalize.
    hour: Number(get("hour")) % 24,
    minute: Number(get("minute")),
    second: Number(get("second")),
    weekday: weekdayMap[get("weekday")],
  };
}

// Naive wall-clock string ("YYYY-MM-DDTHH:mm:ss") of `date` as seen in
// timeZone. This is exactly what Graph expects paired with start.timeZone.
function toGraphDateTime(date: Date, timeZone: string): string {
  const p = getZonedParts(date, timeZone);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${p.year}-${pad(p.month)}-${pad(p.day)}T${pad(p.hour)}:${pad(p.minute)}:${pad(p.second)}`;
}

// Offset (ms) at `date` such that instant = (wall clock read as UTC) - offset.
function timeZoneOffsetMs(date: Date, timeZone: string): number {
  const p = getZonedParts(date, timeZone);
  const wallClockAsUTC = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
  return wallClockAsUTC - date.getTime();
}

// Inverse of toGraphDateTime: interpret a naive wall-clock string (e.g. the
// value from an <input type="datetime-local">) as being in timeZone and return
// the absolute instant it refers to.
function zonedWallClockToInstant(naive: string, timeZone: string): Date {
  // Read the naive components as if they were UTC to get a first guess...
  const guess = new Date(naive.length === 16 ? `${naive}:00Z` : `${naive}Z`);
  // ...then shift by the zone's offset at that instant.
  return new Date(guess.getTime() - timeZoneOffsetMs(guess, timeZone));
}

function buildSearchWindowFromNow(): { start: Date; end: Date } {
  const start = new Date();
  start.setMinutes(0, 0, 0);
  start.setHours(start.getHours() + 1);

  const end = new Date(start);
  end.setDate(end.getDate() + 7);

  return { start, end };
}

app.tab("home", path.join(__dirname, "./client"));

(async () => {
  await app.start(process.env.PORT || 3978);
})();