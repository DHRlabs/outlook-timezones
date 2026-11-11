# Outlook Time Zones

A small Outlook add-in (Office.js, version 0.1.1). Plain static files, no build step.

- Reading: when a message has a clock time with a zone, one thin bar above the body shows it in your zone, like `4pm EST M.Santos | 2pm AZ you`.
- Names in the bar are first initial and last name: Maria Santos shows as `M.Santos`. A one-word name stays whole.
- Writing: the "Add time zones" button rewrites each clock time in your draft as `2pm AZ [3pm CT / 4pm ET]`. Clicking twice changes nothing.
- Your zone is your device zone. Times with no zone written are never guessed.
- When the writer already gave the same time in two zones (`4pm ET / 1pm PT`), the bar shows it once and the draft is left alone. Two different times (`2pm ET / 3pm CT`) are both read and both converted.
- Hosted on GitHub Pages at https://dhrlabs.github.io/outlook-timezones/.
- Needs Outlook on the web, new Outlook, or Outlook for Mac (mailbox requirement set 1.5).

## Time forms it reads

`4pm`, `4 pm`, `4PM`, `4p.m.`, `4p`, `4:30pm`, `4.30pm`, `4pmET`, `4PMEST`, `16:00`, `09:00`, `1600 hrs`,
noon, midnight, ranges (`2-3pm ET`, `2 to 3pm ET`, `between 2 and 3pm ET`), and a weekday beside the time
(`Thu`, `thursday`). The conversion uses the weekday word when one is written next to the time, and the
message date otherwise. Zones can be an abbreviation, a spelled-out name (`Eastern`,
`Eastern Standard Time`), or an exact offset (`UTC-5`, `GMT+1`, `UTC-07:00`).

Deliberately not read: `4 o'clock`, `half past 4`, a bare `4` with no meridiem and no zone, a calendar
date beside the time (`11/12`, `Nov 12`, `tomorrow`), and any abbreviation with more than one meaning
(IST, BST, AST and the rest).

## Sideload in Outlook on the web

1. Open Outlook on the web, open any message, choose the apps button (or More actions, Get Add-ins), then My add-ins.
2. Under Custom add-ins choose Add a custom add-in, then Add from file, and pick `manifest.xml`.
3. Open a message, click the Time Zones button on the ribbon, and pin the pane so the bar updates on every message.
4. In a new message, click Add time zones.

## Known limits

- The bar only appears while the Time Zones pane is open and pinned (Outlook has no event for opening a message).
- No bar on phone.
- Outside senders need to type a zone.
- Coworker zones are planned for 0.2.
