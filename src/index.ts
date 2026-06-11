import "dotenv/config";
import "isomorphic-fetch";

import fs from "fs";
import https from "https";
import path from "path";
import express from "express";
import Database from "better-sqlite3";

import { App, ExpressAdapter, IPlugin } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";

import { ClientSecretCredential } from "@azure/identity";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
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

const db = new Database(path.join(process.cwd(), "coffee-talk.db"));

db.exec(`
  CREATE TABLE IF NOT EXISTS meetings (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    channel_id TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS meeting_participants (
    meeting_id INTEGER NOT NULL REFERENCES meetings(id),
    user_id    TEXT NOT NULL,
    PRIMARY KEY (meeting_id, user_id)
  );
`);

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

  const allowed = members.filter(
    (m) => m.email && ALLOWED_EMAILS.has(m.email.toLowerCase())
  );

  // Attach photos only for the members we keep, so we don't hit Graph for users
  // we're about to discard.
  for (const member of allowed) {
    if (member.userId) {
      member.photo = await getUserPhotoDataUri(graphClient, member.userId);
    }
  }

  return allowed;
}

// Fetches a user's profile photo as a data URI ("data:image/jpeg;base64,...")
// so the frontend can drop it straight into an <img src>. Returns undefined
// when the user has no photo (Graph returns 404) or the lookup fails.
async function getUserPhotoDataUri(
  graphClient: Client,
  userId: string
): Promise<string | undefined> {
  try {
    const photo = await graphClient
      .api(`/users/${userId}/photo/$value`)
      .responseType(ResponseType.ARRAYBUFFER)
      .get();

    return `data:image/jpeg;base64,${Buffer.from(photo).toString("base64")}`;
  } catch (error: any) {
    if (error?.statusCode !== 404) {
      console.warn(`Could not load photo for user ${userId}`, error);
    }
    return undefined;
  }
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

type Slot = {
  startDateTime: string;
  endDateTime: string;
  availableParticipants: ChannelMember[];
  startMs: number;
  endMs: number;
};

// Slide a `durationMinutes` window across [window.start, window.end) in 30-min
// steps and return the slot that includes the MOST of `participants` (so we
// drop as few people as possible). A slot where the whole group is free is
// taken immediately; otherwise we keep scanning and return the best slot that
// still has at least `minParticipants` free. When `enforceWorkingHours` is true
// the slot must also fall on a weekday between 09:00 and 17:00 in `timeZone`.
// Slots overlapping any interval in `bookedIntervals` (meetings already placed
// this run) are skipped, so separate groups spread out instead of stacking on
// the same time.
async function findAvailableSlot(
  graphClient: Client,
  participants: ChannelMember[],
  durationMinutes: number,
  timeZone: string,
  minParticipants: number,
  window: { start: Date; end: Date },
  enforceWorkingHours: boolean,
  bookedIntervals: { start: number; end: number }[]
): Promise<Slot | null> {
  let cursor = new Date(window.start);
  let best: Slot | null = null;

  while (cursor < window.end) {
    const slotEnd = addMinutes(cursor, durationMinutes);

    // The whole meeting must fit inside the window.
    if (slotEnd > window.end) {
      break;
    }

    const slotStartMs = cursor.getTime();
    const slotEndMs = slotEnd.getTime();

    let acceptable = true;
    if (enforceWorkingHours) {
      const { weekday, hour } = getZonedParts(cursor, timeZone);
      const isWeekday = weekday !== 0 && weekday !== 6;
      const isWorkingHour = hour >= 9 && hour < 17;
      acceptable = isWeekday && isWorkingHour;
    }

    // Don't reuse a time already taken by a meeting we placed this run.
    if (
      acceptable &&
      bookedIntervals.some((b) => slotStartMs < b.end && slotEndMs > b.start)
    ) {
      acceptable = false;
    }

    if (acceptable) {
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

      // Everyone is free here — can't do better, take it now.
      if (availableParticipants.length === participants.length) {
        return { startDateTime, endDateTime, availableParticipants, startMs: slotStartMs, endMs: slotEndMs };
      }

      // Otherwise remember the slot that keeps the most people (still >= min).
      if (
        availableParticipants.length >= minParticipants &&
        (!best || availableParticipants.length > best.availableParticipants.length)
      ) {
        best = { startDateTime, endDateTime, availableParticipants, startMs: slotStartMs, endMs: slotEndMs };
      }
    }

    cursor = addMinutes(cursor, 30);
  }

  return best;
}

async function createCalendarEvent(
  graphClient: Client,
  participants: ChannelMember[],
  subject: string,
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
  organizerEmail: string
) {
  const attendees = participants.map((p) => ({
    emailAddress: {
      address: p.email,
      name: p.displayName,
    },
    type: "required",
  }));

  const bodyContent = await generateMeetingContent(participants);

  return graphClient
    .api(`/users/${organizerEmail}/events`)
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

    const prompt = profileContent
      ? `Here are the personal profile pages from Confluence for the participants:\n\n${profileContent}\n\nBased on their hobbies and personal interests, generate 4–6 conversation topic suggestions for their coffee meeting. List topics based on shared hobbies or interests first, but do not label or mark them differently from the others — all topics should look the same.\n\nRespond with raw HTML only — no markdown, no code fences, no backticks. Start directly with the first HTML tag:\n<p>Here are some conversation starters for your coffee chat:</p>\n<ul>\n  <li>A topic suggestion...</li>\n  <li>Another topic suggestion...</li>\n</ul>\n<p><em>Enjoy your coffee chat! ☕ — DDS Coffee Talk</em></p>`
      : `Generate 4–6 general conversation topic suggestions for a coffee meeting between ${names}.\n\nRespond with raw HTML only — no markdown, no code fences, no backticks. Start directly with the first HTML tag:\n<p>Here are some conversation starters for your coffee chat:</p>\n<ul>\n  <li>A topic suggestion...</li>\n</ul>\n<p><em>Enjoy your coffee chat! ☕ — DDS Coffee Talk</em></p>`;

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

    const raw = response.output?.message?.content?.find((b: any) => b.text)?.text;
    const text = raw?.replace(/^```[a-z]*\n?/i, "").replace(/\n?```$/i, "").trim();
    if (text) {
      const missingNote =
        missingProfiles.length > 0
          ? `<p style="background:#30414f;color:#ffffff;padding:12px 16px;border-radius:8px;font-family:system-ui,-apple-system,'Segoe UI',sans-serif;font-size:14px;line-height:1.5;margin-top:16px;"><strong>Note:</strong> We noticed that profile pages were not found for ${missingProfiles.join(", ")}. If you're part of this team, please consider creating your personal profile page on Confluence to help your colleagues get to know you better and foster meaningful connections! 👀</p>`
          : "";
      return text + missingNote;
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
      organizer,
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

    const pairs = getPairHistory(channelId);
    const groups = createHistoryAwareGroups(members, min, max, pairs);

    const tz = timeZone || process.env.DEFAULT_TIME_ZONE || "Europe/Madrid";
    const window = buildSearchWindowFromNow();
    const bookedIntervals: { start: number; end: number }[] = [];

    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];
      const slot = await findAvailableSlot(
        graphClient,
        participants,
        Number(durationMinutes),
        tz,
        min,
        window,
        true,
        bookedIntervals
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

      bookedIntervals.push({ start: slot.startMs, end: slot.endMs });

      const event = await createCalendarEvent(
        graphClient,
        slot.availableParticipants,
        `DDS Coffee Talk #${i + 1}`,
        slot.startDateTime,
        slot.endDateTime,
        tz,
        organizer
      );

      recordMeeting(channelId, slot.availableParticipants);

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
      organizer,
      minPerMeeting,
      maxPerMeeting,
      durationMinutes,
      startDateTime,
      endDateTime,
      timeZone,
    } = req.body;

    if (!teamId || !channelId || !startDateTime || !endDateTime) {
      return res.status(400).json({
        ok: false,
        message: "Missing team/channel context or start/end of the range.",
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

    // Inputs are day-only ("YYYY-MM-DD") in the user's tz. The range spans from
    // the start of the first day to the end of the last (start of the day after
    // endDate, exclusive). The 09:00–17:00 working-hour gate inside
    // findAvailableSlot keeps slots within sensible hours of those days.
    const rangeStart = zonedWallClockToInstant(`${startDateTime}T00:00`, tz);
    const rangeEnd = addMinutes(
      zonedWallClockToInstant(`${endDateTime}T00:00`, tz),
      24 * 60
    );

    if (rangeEnd.getTime() <= rangeStart.getTime()) {
      return res.status(400).json({
        ok: false,
        message: "The end day must be on or after the start day.",
        meetings: [],
      });
    }

    // Never search into the past — if the start day is today, begin from now.
    const now = new Date();
    const window = {
      start: rangeStart.getTime() < now.getTime() ? now : rangeStart,
      end: rangeEnd,
    };

    const pairs = getPairHistory(channelId);
    const groups = createHistoryAwareGroups(members, min, max, pairs);

    const bookedIntervals: { start: number; end: number }[] = [];
    const meetings: PlannedMeeting[] = [];

    for (let i = 0; i < groups.length; i++) {
      const participants = groups[i];

      // Search inside the chosen days, restricted to working hours.
      const slot = await findAvailableSlot(
        graphClient,
        participants,
        Number(durationMinutes),
        tz,
        min,
        window,
        true,
        bookedIntervals
      );

      if (!slot) {
        meetings.push({
          subject: `DDS Coffee Talk #${i + 1}`,
          participants,
          status: "failed",
          message: `No slot with at least ${min} participant(s) free in the selected range.`,
        });
        continue;
      }

      bookedIntervals.push({ start: slot.startMs, end: slot.endMs });

      const event = await createCalendarEvent(
        graphClient,
        slot.availableParticipants,
        `DDS Coffee Talk #${i + 1}`,
        slot.startDateTime,
        slot.endDateTime,
        tz,
        organizer
      );

      recordMeeting(channelId, slot.availableParticipants);

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

type ChannelMember = {
  id?: string;
  userId?: string;
  displayName: string;
  email?: string;
  photo?: string; // data URI of the user's profile photo, if any
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

function getPairHistory(channelId: string): Set<string> {
  const rows = db
    .prepare(
      `
    SELECT a.user_id as user1, b.user_id as user2
    FROM meeting_participants a
    JOIN meeting_participants b ON a.meeting_id = b.meeting_id AND a.user_id < b.user_id
    JOIN meetings m ON a.meeting_id = m.id
    WHERE m.channel_id = ?
  `,
    )
    .all(channelId) as { user1: string; user2: string }[];
  return new Set(rows.map((r) => `${r.user1}:${r.user2}`));
}

const insertMeeting = db.prepare(
  `INSERT INTO meetings (channel_id, created_at) VALUES (?, ?)`,
);
const insertParticipant = db.prepare(
  `INSERT INTO meeting_participants (meeting_id, user_id) VALUES (?, ?)`,
);

const recordMeeting = db.transaction(
  (channelId: string, participants: ChannelMember[]) => {
    const { lastInsertRowid } = insertMeeting.run(
      channelId,
      new Date().toISOString(),
    );
    for (const p of participants) {
      if (p.userId) insertParticipant.run(lastInsertRowid, p.userId);
    }
  },
);

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

function countRepeatedPairs(groups: ChannelMember[][], pairHistory: Set<string>): number {
  let count = 0;
  for (const group of groups) {
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const key = [group[i].userId!, group[j].userId!].sort().join(":");
        if (pairHistory.has(key)) count++;
      }
    }
  }
  return count;
}

// Tries 200 random shuffles and returns the partition with the fewest repeated
// pairs. A single greedy pass fails because the insertion order determines which
// members share a group — trying many orderings and keeping the global minimum
// consistently finds the best partition for small team sizes.
function createHistoryAwareGroups(
  members: ChannelMember[],
  minPerMeeting: number,
  maxPerMeeting: number,
  pairHistory: Set<string>,
): ChannelMember[][] {
  const usable = members.filter((m) => m.email && m.userId);
  const groupCount = Math.floor(usable.length / minPerMeeting);
  if (groupCount === 0) return [];

  let bestGroups: ChannelMember[][] = [];
  let bestScore = Infinity;

  for (let attempt = 0; attempt < 200; attempt++) {
    const shuffled = [...usable].sort(() => Math.random() - 0.5);
    const groups: ChannelMember[][] = Array.from({ length: groupCount }, () => []);

    for (const member of shuffled) {
      let bestGroup = -1;
      let bestGroupScore = Infinity;

      for (let g = 0; g < groupCount; g++) {
        if (groups[g].length >= maxPerMeeting) continue;

        const repeatedPairs = groups[g].filter((existing) => {
          const key = [existing.userId!, member.userId!].sort().join(":");
          return pairHistory.has(key);
        }).length;

        const score = groups[g].length * 1000 + repeatedPairs;
        if (score < bestGroupScore) {
          bestGroupScore = score;
          bestGroup = g;
        }
      }

      if (bestGroup !== -1) groups[bestGroup].push(member);
    }

    const validGroups = groups.filter((g) => g.length >= minPerMeeting);
    const score = countRepeatedPairs(validGroups, pairHistory);

    if (score < bestScore) {
      bestScore = score;
      bestGroups = validGroups;
    }

    if (bestScore === 0) break;
  }

  return bestGroups;
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