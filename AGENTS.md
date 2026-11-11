# Outlook Time Zones

- Plain static files: HTML and JavaScript, no build step, no package.json, no dependencies.
- One zone table, in `src/timezones.js`. Parsing and conversion live only there; read and compose both use it.
- Never guess a zone. On the read side, a time with no written zone is skipped.
- Zone invariant: a time that has a zone word next to it (abbreviation or spelled-out name, any supported separator) has exactly two outcomes: it is converted from that zone, or it is left completely untouched. It never gets the reader's or writer's own zone label. Only a time with no zone word next to it gets the writer's label in compose.
- A spelled-out name (Eastern, Central, Mountain, Pacific, Arizona) is a zone when it is followed by "Time" / "Standard Time" / "Daylight Time", ends the clause, is followed by punctuation, or is followed by a function word (or, and, if, with, for, on, at, is, works, but, so, then, to, please, today, tomorrow, tonight). Followed by any other word ("Central Park", "Pacific Coast"), the time is left untouched.
- Tests are end-to-end only (real Outlook or the real shared logic run from the command line). No unit tests, no tests folder.
- American English.
- Version 0.2 backlog: coworker zones via Microsoft Graph getSchedule (needs an app registration).
