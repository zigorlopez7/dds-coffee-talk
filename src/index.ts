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
import { BedrockRuntimeClient, ConverseCommand } from "@aws-sdk/client-bedrock-runtime";
import { fromNodeProviderChain } from "@aws-sdk/credential-providers";

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

async function getAvailableParticipants(
  graphClient: Client,
  participants: ChannelMember[],
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
  durationMinutes: number
): Promise<ChannelMember[]> {
  const caller = participants.find((p) => p.userId) ?? participants[0];

  const availabilityViewInterval = Math.max(5, Math.floor(durationMinutes / 2));

  const result = await graphClient
    .api(`/users/${caller.userId}/calendar/getSchedule`)
    .post({
      schedules: participants.map((p) => p.email),
      startTime: { dateTime: startDateTime, timeZone },
      endTime: { dateTime: endDateTime, timeZone },
      availabilityViewInterval,
    });

  return participants.filter((_, i) => {
    const schedule = result.value[i];
    return !schedule?.scheduleItems || schedule.scheduleItems.length === 0;
  });
}

async function findAvailableSlot(
  graphClient: Client,
  participants: ChannelMember[],
  durationMinutes: number,
  timeZone: string,
  minParticipants: number
): Promise<{ startDateTime: string; endDateTime: string; availableParticipants: ChannelMember[] } | null> {
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

      const availableParticipants = await getAvailableParticipants(
        graphClient,
        participants,
        startDateTime,
        endDateTime,
        timeZone,
        durationMinutes
      );

      if (availableParticipants.length >= minParticipants) {
        return { startDateTime, endDateTime, availableParticipants };
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

  const bodyContent = await generateMeetingContent(participants);

  return graphClient
    .api(`/users/${organizer.userId}/events`)
    .post({
      subject,
      body: {
        contentType: "HTML",
        content: bodyContent,
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

async function fetchParticipantProfiles(
  participants: ChannelMember[],
): Promise<{ content: string; missing: string[] }> {
  const baseUrl =
    process.env.CONFLUENCE_BASE_URL ?? "https://dehngroup.atlassian.net";
  const email = process.env.CONFLUENCE_EMAIL;
  const apiToken = process.env.CONFLUENCE_API_TOKEN;
  const indexPageId = "2839281780";
  const headers = {
    Authorization: `Basic ${Buffer.from(`${email}:${apiToken}`).toString("base64")}`,
    "Content-Type": "application/json",
    Accept: "application/json",
  };

  const childRes = await fetch(
    `${baseUrl}/wiki/rest/api/content/${indexPageId}/child/page?limit=100`,
    { headers },
  );

  if (!childRes.ok) {
    throw new Error(`Confluence API returned ${childRes.status}`);
  }

  const { results: childPages } = (await childRes.json()) as {
    results: { id: string; title: string }[];
  };

  const results = await Promise.all(
    participants.map(async (participant) => {
      const name = participant.displayName.toLowerCase();
      const match = childPages.find((p) => {
        const title = p.title.toLowerCase();
        return (
          title === name ||
          title.includes(name) ||
          name
            .split(" ")
            .some((part) => part.length > 2 && title.includes(part))
        );
      });

      if (!match) return { name: participant.displayName, text: null };

      const pageRes = await fetch(
        `${baseUrl}/wiki/rest/api/content/${match.id}?expand=body.storage`,
        { headers },
      );
      if (!pageRes.ok) return { name: participant.displayName, text: null };

      const pageData = (await pageRes.json()) as {
        body: { storage: { value: string } };
      };
      const text = pageData.body.storage.value
        .replace(/<[^>]+>/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 2000);

      return { name: participant.displayName, text };
    }),
  );

  const content = results
    .filter((r) => r.text)
    .map((r) => `--- ${r.name} ---\n${r.text}`)
    .join("\n\n");

  const missing = results.filter((r) => !r.text).map((r) => r.name);

  return { content, missing };
}

async function generateMeetingContent(
  participants: ChannelMember[],
): Promise<string> {
  const awsRegion = process.env.AWS_REGION ?? "eu-central-1";
  const modelId =
    process.env.BEDROCK_MODEL_ID ??
    "eu.anthropic.claude-haiku-4-5-20251001-v1:0";

  const fallback = "<p>Created by DDS Coffee Talk. Enjoy your coffee chat!</p>";

  const names = participants.map((p) => p.displayName).join(", ");

  try {
    let profileContent = "";
    let missingProfiles: string[] = [];
    try {
      ({ content: profileContent, missing: missingProfiles } =
        await fetchParticipantProfiles(participants));
    } catch (err) {
      console.warn("Could not fetch Confluence profiles:", err);
    }

    const missingNote =
      missingProfiles.length > 0
        ? `\n\nNote: no Confluence profile was found for ${missingProfiles.join(", ")}. At the end of the response, add a short reminder (as a <p> tag) asking them to create their personal profile page.`
        : "";

    const prompt = profileContent
      ? `Here are the personal profile pages from Confluence for the participants:\n\n${profileContent}\n\nBased on their hobbies and personal interests, generate 4–6 conversation topic suggestions for their coffee meeting. List topics based on shared hobbies or interests first (mark them with "[Shared]").${missingNote}\n\nFormat the response as an HTML fragment (no <html>, <head>, or <body> tags):\n<p>Here are some conversation starters for your coffee chat:</p>\n<ul>\n  <li><strong>[Shared]</strong> A topic based on a shared interest...</li>\n  <li>Another topic suggestion...</li>\n</ul>\n<p><em>Enjoy your coffee chat! ☕ — DDS Coffee Talk</em></p>`
      : `Generate 4–6 general conversation topic suggestions for a coffee meeting between ${names}.${missingNote}\n\nFormat the response as an HTML fragment (no <html>, <head>, or <body> tags):\n<p>Here are some conversation starters for your coffee chat:</p>\n<ul>\n  <li>A topic suggestion...</li>\n</ul>\n<p><em>Enjoy your coffee chat! ☕ — DDS Coffee Talk</em></p>`;

    const client = new BedrockRuntimeClient({
      region: awsRegion,
      credentials: fromNodeProviderChain(),
    });

    const response = await client.send(
      new ConverseCommand({
        modelId,
        messages: [{ role: "user", content: [{ text: prompt }] }],
        inferenceConfig: { maxTokens: 1024 },
      }),
    );

    const text = response.output?.message?.content?.find((b) => b.text)?.text;
    if (text) {
      return text;
    }
  } catch (error) {
    console.warn("Could not generate AI meeting content:", error);
  }

  return fallback;
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
      minPerMeeting,
      maxPerMeeting,
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

    const { min, max } = normalizeGroupSize(minPerMeeting, maxPerMeeting);
    const graphClient = await getGraphClient();
    const members = await loadChannelMembers(graphClient, teamId, channelId);

    const groups = createRandomGroups(members, min, max);

    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];
      const slot = await findAvailableSlot(
        graphClient,
        participants,
        Number(durationMinutes),
        timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid",
        min
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
        slot.availableParticipants,
        `DDS Coffee Talk #${i + 1}`,
        slot.startDateTime,
        slot.endDateTime,
        timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid"
      );

      meetings.push({
        subject: event.subject,
        participants: slot.availableParticipants,
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
      minPerMeeting,
      maxPerMeeting,
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

    const { min, max } = normalizeGroupSize(minPerMeeting, maxPerMeeting);
    const tz = timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid";
    const graphClient = await getGraphClient();
    const members = await loadChannelMembers(graphClient, teamId, channelId);

    const groups = createRandomGroups(members, min, max);

    // The picked value (e.g. from <input type="datetime-local">) is a naive
    // wall clock the user means in `tz` — interpret it there, not server-local.
    const start = zonedWallClockToInstant(startDateTime, tz);
    const end = addMinutes(start, Number(durationMinutes));

    const startGraph = toGraphDateTime(start, tz);
    const endGraph = toGraphDateTime(end, tz);

    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];

      const availableParticipants = await getAvailableParticipants(
        graphClient,
        participants,
        startGraph,
        endGraph,
        tz,
        Number(durationMinutes)
      );

      if (availableParticipants.length < min) {
        meetings.push({
          subject: `DDS Coffee Talk #${i + 1}`,
          participants,
          startDateTime: startGraph,
          endDateTime: endGraph,
          status: "failed",
          message: `Only ${availableParticipants.length} participant(s) available — need at least ${min}.`,
        });
        continue;
      }

      const event = await createCalendarEvent(
        graphClient,
        availableParticipants,
        `DDS Coffee Talk #${i + 1}`,
        startGraph,
        endGraph,
        tz
      );

      meetings.push({
        subject: event.subject,
        participants: availableParticipants,
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

// Coerce the request's min/max group size into a valid range: min is at least 1
// and max is never below min. Missing values fall back to pairs.
function normalizeGroupSize(
  minPerMeeting: unknown,
  maxPerMeeting: unknown
): { min: number; max: number } {
  const min = Math.max(1, Number(minPerMeeting) || 2);
  const max = Math.max(min, Number(maxPerMeeting) || min);
  return { min, max };
}

// Partition ALL usable members into as many meeting groups as possible, each
// holding between minPerMeeting and maxPerMeeting people. We open the most
// groups the minimum allows (floor(total / min)) and round-robin everyone into
// them, capping each at the maximum. With evenly divisible numbers every group
// gets the same size; otherwise the extra people are spread one per group up to
// the cap. Anyone who can't fill a group of `min` is left out for this round.
function createRandomGroups(
  members: ChannelMember[],
  minPerMeeting: number,
  maxPerMeeting: number
): ChannelMember[][] {
  const usableMembers = members.filter((m) => m.email);
  const shuffled = shuffle(usableMembers);

  const groupCount = Math.floor(shuffled.length / minPerMeeting);
  if (groupCount === 0) {
    return [];
  }

  const groups: ChannelMember[][] = Array.from({ length: groupCount }, () => []);

  let cursor = 0;
  for (const member of shuffled) {
    // Skip groups that already hit the maximum; stop once they are all full.
    let scanned = 0;
    while (scanned < groupCount && groups[cursor].length >= maxPerMeeting) {
      cursor = (cursor + 1) % groupCount;
      scanned++;
    }
    if (groups[cursor].length >= maxPerMeeting) {
      break;
    }

    groups[cursor].push(member);
    cursor = (cursor + 1) % groupCount;
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