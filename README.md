# SBC Ticket System

A ticket reservation system for shows, built on Google Workspace tools (Google Forms, Sheets, Drive and Gmail) and automated with Google Apps Script. No server or paid service is needed.

Attendees reserve tickets through a Google Form. The system then tracks reservations, sends payment instructions, handles waitlists, and issues ticket PDFs with QR codes once payment is confirmed.

## Features

- One-run setup of the Drive folders, Google Form and response spreadsheet for each show
- Payment options you can switch on or off per show: Bank Transfer, Maya, GCash and PayPal
- Automatic reservation IDs, pricing and waitlisting when the ticket limit is reached
- Branded HTML emails: reservation received, waitlisted, ticket confirmed (with PDF and QR code), cancellation and reminder
- Early-bird pricing that ends at a ticket count or a date/time, whichever comes first, and a configurable sender name
- Global images shared by all shows, with per-show overrides
- A **References** tab that stores every link and setting, so the two scripts stay in sync

## How it works

The project has two Google Apps Script files:

| Script | Where it runs | Purpose |
|--------|---------------|---------|
| `Create_Ticket_Form.gs` | Standalone project at [script.google.com](https://script.google.com) | Run once per show. Creates the folders, the form, the response spreadsheet and the References tab. Downloads the email templates and default images from this repository. |
| `Code_Ticket_Sheet.gs` | Bound to the response spreadsheet (Extensions > Apps Script) | Builds the Tracking, Tickets and Counters tabs, loads the configuration and runs the timed triggers that process reservations and send emails. |

## Repository contents

| Path | Description |
|------|-------------|
| `email_html_templates/` | HTML email templates (downloaded automatically by `Create_Ticket_Form.gs`) |
| `global_assets/` | Default images shared by all shows (downloaded automatically by `Create_Ticket_Form.gs`) |
| [Releases](https://github.com/banting-sbc/SBC-ticket-system/releases) | The two script files, ready to download |

## Quick start

1. (Optional) Create a folder named `SBC-ticket-system` in your Google Drive (My Drive). The setup script creates it if it does not exist.
2. Download `Create_Ticket_Form.gs` and `Code_Ticket_Sheet.gs` from the [latest release](https://github.com/banting-sbc/SBC-ticket-system/releases). If the scripts are not available in the releases section, reach out to the repo owner.
3. At [script.google.com](https://script.google.com), create a new project (name it after the show), paste in `Create_Ticket_Form.gs`, fill in the input section at the top, and run `createReservationForm` function. If a download fails because of GitHub API limits, run it again.
4. Optionally add your own images to `global_assets` (whole group) or `show_assets` (one show). Images in `show_assets` take precedence. File names must match exactly, for example `header`, `banner`, `bank-info`, `bank-qr`, with a `.png` or `.jpg` extension.
5. Open the spreadsheet created by the first script, go to **Extensions > Apps Script**, paste in the contents of the 2nd script, `Code_Ticket_Sheet.gs`, save, then run the `setupSheets` function followed by `createAllTriggers` function.

For the full walkthrough, see the [Setup Guide](https://github.com/banting-sbc/SBC-ticket-system/blob/main/documentation/SBC_Ticket_System_Setup_Guide.md). Once the system is running, see [How to Use](https://github.com/banting-sbc/SBC-ticket-system/blob/main/documentation/SBC_Ticket_System_How_to_Use.md).

## Documentation

| Guide | For | Contents |
|-------|-----|----------|
| [Setup Guide](https://github.com/banting-sbc/SBC-ticket-system/blob/main/documentation/SBC_Ticket_System_Setup_Guide.md) | Show admins | Requirements, required images, creating the form and spreadsheet, installing the scripts and triggers, troubleshooting |
| [How to Use](https://github.com/banting-sbc/SBC-ticket-system/blob/main/documentation/SBC_Ticket_System_How_to_Use.md) | Show admins | The reservation process from responder link to tickets, statuses, ticket limit and waitlist, early bird pricing, timing |

Both guides are also available as Word (`.docx`) files.

## Requirements

- A Google account. A free account works, but Apps Script quotas are much lower than on Google Workspace (for example 100 vs 1,500 email recipients per day), so Workspace is recommended for larger shows. See the [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas).
- Internet access to this public repository, because the setup script downloads the email templates and default images from it.

## Notes

- Unauthenticated GitHub API requests are limited to 60 per hour. Re-running the setup script only downloads files that are still missing.
- The scripts ask for permission to use Drive, Forms, Spreadsheets, Gmail and external requests the first time they run.
