# Demo venue: The Foundry on Ninth

One coherent, fictional venue with example text for every box in the
console, cross-checked for consistency (room names, coordinator names,
and the fact/behavior split between Availability policy and Rules all
agree with each other).

## To load it

1. Open the tool locally (not the production Sheet-synced site).
2. Open the browser's dev console.
3. Paste the contents of `foundry-on-ninth-demo.js` and press enter.
4. Reload the page. A "The Foundry on Ninth (demo)" account appears
   with one location, every field filled in.

Uses the console's own `createAccount` / `createVenue` / `setVenueField`
functions -- the same path a person typing into the form goes through --
so it stamps and saves exactly like a real save would.

Safe to run against a local, unsynced instance only. If a Google Sheet
sync URL is configured, this account will be pushed to it like any other
save.
