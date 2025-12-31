# Christmas Party Invite Automation

This project is a simple email automation built using **Google Apps Script** and **Google Sheets**.

## Overview
The script reads attendee data (name, email, and status) from a Google Sheet and sends a personalized HTML email invite to selected recipients.

## How it Works
- Attendee details are stored in a Google Sheet.
- The Apps Script (`code.js`) is bound to the sheet.
- The script checks a status column (e.g., "Yes").
- If the status is "Yes", an email invite is sent.
- After sending, the status is updated to "Sent" to avoid duplicates.

## Files
- `code.js` – Main Google Apps Script file containing the email logic and HTML template.
- `attendees.csv` – Sample/exported version of the Google Sheet structure (no real data).

## Features
- Personalized HTML email invites
- Organizer section with images
- Google Calendar "Add to Calendar" link
- Google Maps location link
- Status update in Google Sheet after sending

## Notes
This repository contains only the script and sample data.  
The actual Google Sheet and live Apps Script project run within Google Workspace.

