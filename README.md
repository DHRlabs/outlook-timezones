# Outlook Time Zones

A small Outlook add-in (Office.js, version 0.1). Plain static files, no build step.

- Reading: when a message has a clock time with a zone, one thin bar above the body shows it in your zone, like `4pm EST Maria | 2pm AZ you`.
- Writing: the "Add time zones" button rewrites each clock time in your draft as `2pm AZ [3pm CT / 4pm ET]`. Clicking twice changes nothing.
- Your zone is your device zone. Times with no zone written are never guessed.
- Hosted on GitHub Pages at https://dhrlabs.github.io/outlook-timezones/.
- Needs Outlook on the web, new Outlook, or Outlook for Mac (mailbox requirement set 1.5).

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
