Midlothian Friends — Self‑Registration Web App (Revisions)
----------------------------------------------------------
Files:
 • MyRegistrationMarkup.html  — responsive form (max 800px, 100% on phones), site‑style selectors, logo, Save & Cancel.
 • MyRegistrationScript.js    — Apps Script server; now requires last name.
 • logo.png                   — small version used at 90px wide in the form header.

Notes:
 • The HTML uses classes like .content-box and .entry-title to harmonize with the site's style.css when embedded.
 • Client-side validation now requires First and Last name; server-side requires Last name.
 • The Cancel button clears the form and shows a non-submitting status message.
