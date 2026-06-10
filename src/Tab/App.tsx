import React, { useEffect, useState } from "react";
import * as teamsJs from "@microsoft/teams-js";

import "./App.css";

type Member = {
  id?: string;
  userId?: string;
  displayName: string;
  email?: string;
};

type Meeting = {
  subject: string;
  participants: Member[];
  startDateTime?: string;
  endDateTime?: string;
  status: string;
  message: string;
  webLink?: string;
};

export default function App() {
  const [teamsContext, setTeamsContext] = useState<teamsJs.app.Context | null>(null);
  const [message, setMessage] = useState("Loading DDS Coffee Talk...");
  const [members, setMembers] = useState<Member[]>([]);
  const [meetings, setMeetings] = useState<Meeting[]>([]);

  const [meetingCount, setMeetingCount] = useState(1);
  const [participantsPerMeeting, setParticipantsPerMeeting] = useState(2);
  const [durationMinutes, setDurationMinutes] = useState(30);
  const [startDateTime, setStartDateTime] = useState("");
  const [timeZone, setTimeZone] = useState("Europe/Madrid");

  useEffect(() => {
    async function init() {
      try {
        await teamsJs.app.initialize();
        const context = await teamsJs.app.getContext();

        setTeamsContext(context);
        setMessage(`Salute. DDS Coffee Talk is running in ${context.app.host.name}.`);
      } catch (error) {
        console.error(error);
        setMessage("Salute. Running outside Teams or Teams SDK failed.");
      }
    }

    init();
  }, []);

  function getTeamChannelPayload() {
    return {
      teamId: teamsContext?.team?.groupId,
      channelId: teamsContext?.channel?.id,
    };
  }

  async function loadChannelMembers() {
    try {
      setMessage("Checking Graph access and loading channel users...");
      setMembers([]);

      const payload = getTeamChannelPayload();

      const response = await fetch("/api/channel-members", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const data = await response.json();

      setMessage(data.message);
      setMembers(data.members || []);
    } catch (error) {
      console.error(error);
      setMessage("Could not load channel users. Check console.");
    }
  }

  async function createMeetingsNow() {
    try {
      setMessage("Preparing random meetings based on availability...");
      setMeetings([]);

      const payload = {
        ...getTeamChannelPayload(),
        meetingCount,
        participantsPerMeeting,
        durationMinutes,
        timeZone,
      };

      const response = await fetch("/api/random-meetings/now", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const data = await response.json();

      setMessage(data.message);
      setMeetings(data.meetings || []);
    } catch (error) {
      console.error(error);
      setMessage("Could not create meetings. Check console.");
    }
  }

  async function createMeetingsAtTime() {
    try {
      setMessage("Preparing random meetings at the selected time...");
      setMeetings([]);

      const payload = {
        ...getTeamChannelPayload(),
        meetingCount,
        participantsPerMeeting,
        durationMinutes,
        startDateTime,
        timeZone,
      };

      const response = await fetch("/api/random-meetings/at-time", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const data = await response.json();

      setMessage(data.message);
      setMeetings(data.meetings || []);
    } catch (error) {
      console.error(error);
      setMessage("Could not create meetings at the selected time. Check console.");
    }
  }

  return (
    <div className="appShell">
      <main className="card">
        <header className="hero">
          <div>
            <p className="eyebrow">
              DEHN Internal Tool
            </p>

            <h1>DDS Coffee Talk</h1>

            <p className="subtitle">
              Create random coffee meetings between channel members
              based on Microsoft Teams and Outlook availability.
            </p>
          </div>
        </header>

        <section className="statusBox">
          <strong>Status</strong>
          <p>{message}</p>
        </section>

        <section className="section">
          <h2>Channel users</h2>
          <p className="muted">
            This checks whether Graph access is configured. Until then, it will show the pending message.
          </p>

          <button className="primaryButton" onClick={loadChannelMembers}>
            Show channel users
          </button>

          {members.length > 0 && (
            <ul className="list">
              {members.map((member, index) => (
                <li key={member.userId || member.id || index}>
                  <strong>{member.displayName}</strong>
                  {member.email && <span>{member.email}</span>}
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className="section">
          <h2>Meeting setup</h2>

          <div className="grid">
            <label>
              Number of meetings
              <input
                type="number"
                min={1}
                value={meetingCount}
                onChange={(e) => setMeetingCount(Number(e.target.value))}
              />
            </label>

            <label>
              Participants per meeting
              <input
                type="number"
                min={2}
                value={participantsPerMeeting}
                onChange={(e) => setParticipantsPerMeeting(Number(e.target.value))}
              />
            </label>

            <label>
              Duration in minutes
              <input
                type="number"
                min={15}
                step={15}
                value={durationMinutes}
                onChange={(e) => setDurationMinutes(Number(e.target.value))}
              />
            </label>

            <label>
              Time zone
              <input
                type="text"
                value={timeZone}
                onChange={(e) => setTimeZone(e.target.value)}
              />
            </label>
          </div>

          <div className="actions">
            <button className="primaryButton" onClick={createMeetingsNow}>
              Create random meetings now
            </button>
          </div>
        </section>

        <section className="section">
          <h2>Schedule at a specified time</h2>

          <label>
            Start date/time
            <input
              type="datetime-local"
              value={startDateTime}
              onChange={(e) => setStartDateTime(e.target.value)}
            />
          </label>

          <div className="actions">
            <button
              className="secondaryButton"
              onClick={createMeetingsAtTime}
              disabled={!startDateTime}
            >
              Create random meetings at selected time
            </button>
          </div>
        </section>

        {meetings.length > 0 && (
          <section className="section">
            <h2>Meeting results</h2>

            <div className="meetingList">
              {meetings.map((meeting, index) => (
                <article className="meetingCard" key={index}>
                  <div>
                    <strong>{meeting.subject}</strong>
                    <span className={`pill ${meeting.status}`}>
                      {meeting.status}
                    </span>
                  </div>

                  <p>{meeting.message}</p>

                  {meeting.startDateTime && (
                    <p className="muted">
                      {meeting.startDateTime} → {meeting.endDateTime}
                    </p>
                  )}

                  <ul>
                    {meeting.participants.map((participant, i) => (
                      <li key={participant.userId || participant.email || i}>
                        {participant.displayName}
                      </li>
                    ))}
                  </ul>

                  {meeting.webLink && (
                    <a href={meeting.webLink} target="_blank" rel="noreferrer">
                      Open meeting
                    </a>
                  )}
                </article>
              ))}
            </div>
          </section>
        )}

        <section className="debugBox">
          <h3>Debug context</h3>
          <pre>
            {JSON.stringify(
              {
                teamId: teamsContext?.team?.groupId,
                channelId: teamsContext?.channel?.id,
                user: teamsContext?.user?.userPrincipalName,
              },
              null,
              2
            )}
          </pre>
        </section>
      </main>
    </div>
  );
}