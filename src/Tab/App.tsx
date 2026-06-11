import { useEffect, useState } from "react";
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

  const [minPerMeeting, setMinPerMeeting] = useState(2);
  const [maxPerMeeting, setMaxPerMeeting] = useState(3);
  const [durationMinutes, setDurationMinutes] = useState(30);
  const [startDateTime, setStartDateTime] = useState("");
  const [endDateTime, setEndDateTime] = useState("");
  const [timeZone, setTimeZone] = useState("Europe/Madrid");

  const [loadingMembers, setLoadingMembers] = useState(false);
  const [loadingNow, setLoadingNow] = useState(false);
  const [loadingAtTime, setLoadingAtTime] = useState(false);

  function formatDateTime(dateTimeString?: string) {
    if (!dateTimeString) return "";

    const date = new Date(dateTimeString);

    return new Intl.DateTimeFormat("en-GB", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
      timeZone: timeZone,
    }).format(date);
  }

  useEffect(() => {
    async function init() {
      try {
        await teamsJs.app.initialize();
        const context = await teamsJs.app.getContext();

        if (context.page.frameContext === teamsJs.FrameContexts.settings) {
          teamsJs.pages.config.setValidityState(true);
          teamsJs.pages.config.registerOnSaveHandler((saveEvent) => {
            teamsJs.pages.config.setConfig({
              suggestedDisplayName: "DDS Coffee Talk",
              contentUrl: `${window.location.origin}/tabs/home`,
              websiteUrl: `${window.location.origin}/tabs/home`,
            });
            saveEvent.notifySuccess();
          });
          setMessage("Click Save to add DDS Coffee Talk to this channel.");
          return;
        }

        setTeamsContext(context);
        setMessage('');
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
      setLoadingMembers(true);
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
    } finally {
      setLoadingMembers(false);
    }
  }

  async function createMeetingsNow() {
    try {
      setLoadingNow(true);
      setMessage("Preparing random meetings based on availability...");
      setMeetings([]);

      const payload = {
        ...getTeamChannelPayload(),
        organizer: teamsContext?.user?.userPrincipalName,
        minPerMeeting,
        maxPerMeeting,
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
    } finally {
      setLoadingNow(false);
    }
  }

  async function createMeetingsAtTime() {
    try {
      setLoadingAtTime(true);
      setMessage("Preparing random meetings within the selected day range...");
      setMeetings([]);

      const payload = {
        ...getTeamChannelPayload(),
        organizer: teamsContext?.user?.userPrincipalName,
        minPerMeeting,
        maxPerMeeting,
        durationMinutes,
        startDateTime,
        endDateTime,
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
    } finally {
      setLoadingAtTime(false);
    }
  }

  return (
    <div className="appShell">
      <main className="card">
        <header className="hero">
          <div>
           
            <h1>DDS Coffee Talk</h1>

            <p className="subtitle">
              Create random coffee meetings between channel members
              based on Microsoft Teams and Outlook availability.
            </p>
          </div>
        </header>

        {message && (
          <section className="statusBox">
            <strong>Status</strong>
            <p>{message}</p>
          </section>
        )}

        <section className="section">
          <h2>Channel users</h2>
        
          <button className="primaryButton" onClick={loadChannelMembers} disabled={loadingMembers}>
            {loadingMembers && <span className="spinner" />}
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
              Min participants per meeting
              <input
                type="number"
                min={1}
                value={minPerMeeting}
                onChange={(e) => setMinPerMeeting(Number(e.target.value))}
              />
            </label>

            <label>
              Max participants per meeting
              <input
                type="number"
                min={minPerMeeting}
                value={maxPerMeeting}
                onChange={(e) => setMaxPerMeeting(Number(e.target.value))}
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
       
          </div>

          <div className="grid">
            <label>
              From day
              <input
                type="date"
                value={startDateTime}
                onChange={(e) => setStartDateTime(e.target.value)}
              />
            </label>

            <label>
              To day
              <input
                type="date"
                min={startDateTime || undefined}
                value={endDateTime}
                onChange={(e) => setEndDateTime(e.target.value)}
              />
            </label>
          </div>

          <div className="actions">
            <button className="primaryButton" onClick={createMeetingsNow} disabled={loadingNow}>
              {loadingNow && <span className="spinner" />}
              Create random meetings now
            </button>
            <button
              className="secondaryButton"
              onClick={createMeetingsAtTime}
              disabled={
                !startDateTime ||
                !endDateTime ||
                endDateTime < startDateTime ||
                loadingAtTime
              }
            >
              {loadingAtTime && <span className="spinner" />}
              Create random meetings in range
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
                      {formatDateTime(meeting.startDateTime)} → {formatDateTime(meeting.endDateTime)}
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
      </main>
    </div>
  );
}