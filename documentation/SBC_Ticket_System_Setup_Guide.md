# SBC Ticket System

## Setup Documentation

This guide explains how to set up the ticket reservation system for a show. It uses two Google Apps Script files:

- `Create_Ticket_Form.gs` – run once per show. It creates the Drive folders, the reservation Google Form, the response spreadsheet and a **References** tab that stores every link and setting.

- `Code_Ticket_Sheet.gs` – added to the response spreadsheet. It builds the working tabs, loads the configuration and runs the timed jobs that process reservations and send the emails.

# 1. Requirements

## 1.1 Google account

A Google account is required. A **free Google account works**, but the daily limits of Google Apps Script are much lower than on Google Workspace:

| **Limit (per day)**             | **Free (gmail.com) account** | **Google Workspace account** |
|---------------------------------|------------------------------|------------------------------|
| Email recipients                | 100                          | 1,500                        |
| Total runtime of timed triggers | 90 minutes                   | 6 hours                      |

Every reservation produces several emails (reservation received, confirmation, and possibly waitlist, reminder or cancellation). On a free account the email limit can be reached quickly for larger shows, and emails will stop being sent until the limit resets. Google Workspace is recommended for shows with many reservations.

Figures are taken from Google's Apps Script quotas page and may change: https://developers.google.com/apps-script/guides/services/quotas

## 1.2 Other prerequisites

- A folder named **SBC-ticket-system** in the **My Drive** of the account that will run the scripts. If it does not exist, the setup script creates it for you.

- The two script files, `Create_Ticket_Form.gs` and `Code_Ticket_Sheet.gs`, downloaded from the GitHub releases page (see "Before you start" in section 2).

- Internet access to the public GitHub repository banting-sbc/`SBC-ticket-system`. The script downloads the HTML email templates and the default global images from it.

## 1.3 Required images

The automation uses the images below. Create them for your group and show, using these **exact file names** (lowercase, with hyphens). Files can be **.png or .jpg**.

| **Image**                 | **File name** | **Recommended size**          | **Purpose**                                                     |
|---------------------------|---------------|-------------------------------|-----------------------------------------------------------------|
| Header image              | header        | 900 x 900 px                  | Group / production logo.                                        |
| Banner image              | banner        | 900 x 450 px                  | Banner for the show / concert.                                  |
| Bank transfer infographic | bank-info     | 1080 x 1200 px                | Payment instructions for bank transfers, in infographic format. |
| Maya infographic          | maya-info     | 1080 x 1200 px                | Payment instructions for Maya payments, in infographic format.  |
| GCash infographic         | gcash-info    | 1080 x 1200 px                | Payment instructions for GCash payments, in infographic format. |
| Bank QR                   | bank-qr       |      | QR code for bank transfers.                                     |
| Maya QR                   | maya-qr       |      | QR code for Maya payments.                                      |
| GCash QR                  | gcash-qr      |      | QR code for GCash payments.                                     |

The script also requires three social icons – facebook, instagram and twitter – which are shown in the email footer. Default versions are downloaded automatically from GitHub, so you only need to supply your own if you want different icons. **[Recommended size: 80 x 80 px]**

<img width="722" height="1086" alt="email-template-updated" src="https://github.com/user-attachments/assets/db15a458-fca1-4663-8ab4-8c36579f6f7c" />

## 1.4 How the asset folders work

Images are looked up in two Google Drive folders:

- **global_assets** (inside `SBC-ticket-system`) – shared by all shows. It is the **default source** of images for the automation. The script creates it automatically and fills it with the default images from GitHub.

- **show_assets** (inside each show's folder) – specific to one show. It is created empty. **Any image placed here takes precedence** over the image with the same name in `global_assets`.

The choice is made image by image. For example, a show can have its own banner in `show_assets` while still using the group's header from `global_assets`.

> **Important:** File names are strictly required. The automation finds images **only by name** (header, banner, bank-info, bank-qr, and so on, with a .png or .jpg extension). An image with any other name is ignored, and the matching link in the spreadsheet stays empty.

# 2. Process

The setup has three steps, to be done in this order, after downloading the scripts.

## Before you start – Download the scripts

1.  Open the releases page: **https://github.com/banting-sbc/`SBC-ticket-system`/releases**.

2.  Select the latest release (currently **v1.0.0**).

3.  Under **Assets**, download `Create_Ticket_Form.gs` and `Code_Ticket_Sheet.gs`.

4.  If the releases are not available, reach out to the repo owner.

5.  Open each file in a text editor so you can copy its full contents later. Do not edit the code; only the input section of `Create_Ticket_Form.gs` is changed in Step 1.

Use the latest release each time you set up a new show, so you get the most recent fixes.

## Step 1 – Create the show form and folders (`Create_Ticket_Form.gs`)

1.  Go to **script.google.com** and click **New project**.

2.  Rename the project. The recommended name is the **show name**.

3.  Open `Create_Ticket_Form.gs`, copy the whole script and paste it into the project's code editor, replacing the default content. Save the project.

4.  Update the fields in the **input section** at the top of the script (table below).

5.  In the function dropdown at the top of the editor, select **createReservationForm** and click **Run**.

6.  On the first run, Google asks you to authorize the script (Drive, Forms, Spreadsheets and external requests). Review the permissions and allow them. For an unverified script, choose **Advanced** and then **Go to project**.

7.  When the run finishes, open the execution log to see the links to the new folders, the form and the spreadsheet.

### Input section

| **Field**                                                 | **Description**                                                                                                                          | **Example**                       |
|-----------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|-----------------------------------|
| `showName` | Name of the show. Used for the form, the spreadsheet and the show folder.                                                                | 'Summer Sessions'                 |
| `showPrefix` | Prefix of the reservation IDs.                                                                                                           | 'SBC'                             |
| `ticketPriceOnline` | Online ticket price (number).                                                                                                            | 500                               |
| `ticketPriceDoor` | Ticket price at the door (number).                                                                                                       | 600                               |
| `ticketLimit` | Total number of tickets (number). Reservations beyond it are waitlisted.                                                                 | 150                               |
| `bankName` | Name of the bank. If filled in, the payment option reads "Bank Transfer (Bank Name)".                                                    | 'BDO'                             |
| `bankAccountName` | Name on the bank account.                                                                                                                | 'Juan Dela Cruz'                  |
| `bankAccountNumber` | Bank account number. Keep it in quotes so leading zeros are kept.                                                                        | '001234567890'                    |
| `gcashNumber` | GCash mobile number. Keep it in quotes.                                                                                                  | '09171234567'                     |
| `paypalLink` | PayPal payment link.                                                                                                                     | 'https://paypal.me/…'             |
| `paypalTransactionFee` | PayPal transaction fee per ticket (number).                                                                                              | 25                                |
| `showDate` | Date of the show, in the format Month dd, YYYY.                                                                                          | 'July 12, 2025'                   |
| `facebookUrl`, `twitterUrl`, `instagramUrl` | Profile links for the email footer icons.                                                                                                | 'https://facebook.com/…'          |
| `emailSenderName` | Sender name shown on the emails.                                                                                                         | 'Sleeping Boy Collective Tickets' |
| `earlyBirdEnabled` | Turns early bird pricing on: 'Yes' or 'No'. Anything other than 'Yes' is treated as 'No'. If 'Yes', the three fields below are required. | 'No'                              |
| `earlyBirdTicketLimit` | Number of tickets sold at the early bird price (whole number).                                                                           | 50                                |
| `earlyBirdDateTimeLimit` | Last moment the early bird price applies. Format YYYY-MM-DD HH:mm, 24-hour clock, in your own time zone.                                 | '2026-12-01 23:59'                |
| `earlyBirdTicketPrice` | Early bird ticket price (number).                                                                                                        | 400                               |
| `enableBankTransfer`, `enableMaya`, `enableGCash`, `enablePaypal` | Payment options to offer in the form: 'Yes' or 'No'. At least one must be Yes.                                                           | 'Yes' / 'No'                      |

Fields still showing a placeholder such as \<BANK NAME HERE\> are ignored when the spreadsheet is filled in later.

### What the script does

1.  Checks the inputs. It stops with a message if a payment option value is invalid or none is enabled, or if early bird pricing is set to Yes but its limit, date/time or price is missing or invalid.

2.  Looks for the **SBC-ticket-system** folder in My Drive and **creates it if missing**.

3.  Looks for the **email_html_templates** folder inside it and **creates it if missing**. It then checks for the eight email templates and **downloads any missing ones from GitHub**: Reminder.html, ReservationCancelled.html, ReservationReceived_BANK.html, ReservationReceived_GCash.html, ReservationReceived_Paypal.html, TicketConfirmed.html, TicketPDF.html and Waitlisted.html.

4.  Looks for the **global_assets** folder and **creates it if missing**. It then checks for the global images (bank-info, bank-qr, banner, facebook, gcash-info, gcash-qr, header, instagram, maya-info, maya-qr, twitter; .png or .jpg) and **downloads any missing ones from GitHub**.

5.  Stops if a folder with the same show name already exists inside `SBC-ticket-system`, so that nothing is overwritten.

6.  Creates the **Google Form** with the reservation questions and the payment options you enabled, and the linked **response spreadsheet**. The spreadsheet uses the time zone of the account running the script, so response timestamps are in local time; the time zone is saved in the References tab.

7.  Creates the **show folder** inside `SBC-ticket-system`, with a QR subfolder (generated ticket QR codes and PDFs) and an empty `show_assets` subfolder, and moves the form and spreadsheet into it.

8.  Adds a **References** tab to the spreadsheet containing all folder links, form links, the show settings and the template file IDs.

The resulting structure in Google Drive:

```
My Drive
└── `SBC-ticket-system`
    ├── `email_html_templates`   (8 HTML email templates)
    ├── `global_assets`          (default images for all shows)
    └── <Show Name>
        ├── <Show Name>        (Google Form)
        ├── <Show Name>        (response spreadsheet)
        ├── QR
        └── `show_assets`        (show-specific images, empty at first)
```

> **If a download is interrupted:** GitHub limits the number of requests from unauthenticated users, so downloads can occasionally fail. If that happens, simply **run the script again**. It only downloads files that are still missing and never overwrites existing ones. Nothing for the show is created until all templates and images are in place.

## Step 2 – Add your custom images

1.  In Google Drive, open `SBC-ticket-system`. Upload images for the whole group to **global_assets**, or images for this show only to the show's **show_assets** folder.

2.  Use the exact file names from section 1.3. Images in `show_assets` override same-named images in `global_assets`.

3.  To replace a default image in `global_assets` (for example the placeholder header), make sure only one file with that name remains – delete or overwrite the old one.

> **Do this before Step 3.** The spreadsheet setup reads the images when it runs. If you add or change images afterwards, run `setupSheets` again: it fills in image links that are still empty, but it does not overwrite links that already have a value.

## Step 3 – Set up the spreadsheet (`Code_Ticket_Sheet.gs`)

1.  Open the response spreadsheet that was created in the show folder.

2.  Click **Extensions \> Apps Script**. A new, untitled script project connected to the spreadsheet opens.

3.  Rename the project to an appropriate name (for example the show name).

4.  Open `Code_Ticket_Sheet.gs`, copy the whole script and paste it into the project's code editor, replacing the default content. **Save** the project.

5.  In the function dropdown, select **setupSheets** and click **Run**. Authorize the script when asked (Spreadsheets, Drive, Gmail and external requests).

6.  When it has finished, check the execution log for warnings. Then select **createAllTriggers** from the dropdown and click **Run**.

### What `setupSheets` does

- Creates the **Tracking**, **Tickets** and **Counters** tabs.

- Fills the **Counters** tab with the show settings from the References tab: show name and prefix, ticket price and limit, bank details, GCash number, PayPal link and fee, show date, social links, email sender name, the early bird settings and the QR folder ID. It also stores the email template IDs.

- Looks up the images in `show_assets` and then `global_assets`, and writes their links to the Counters tab (header, banner, bank / Maya / GCash QR and infographic, and the social icons).

- Sets those images to "Anyone with the link can view" so they can be shown in emails.

- Locks the counter cells that are filled by the script or by formulas, so they cannot be edited by hand.

Values you type into the Counters tab yourself are never overwritten when the function is run again.

### What `createAllTriggers` does

It creates the timed triggers that run the automation. It is safe to run again: triggers that already exist are skipped.

| **Function**          | **Runs**         | **Purpose**                                                                               |
|-----------------------|------------------|-------------------------------------------------------------------------------------------|
| `setReservation` | Every 15 minutes | Assigns reservation IDs, prices and statuses; waitlists when the ticket limit is reached. |
| `sendReservationEmails` | Every 30 minutes | Sends the "reservation received" email with the payment instructions.                     |
| `sendWaitlistedEmails` | Every 5 minutes  | Sends the waitlist notification.                                                          |
| `sendConfirmationEmail` | Every hour       | Generates the ticket PDFs and sends the confirmation once payment is confirmed.           |
| `sendCancellationEmail` | Every hour       | Sends the cancellation email for cancelled reservations.                                  |

The reminder email trigger (`sendReminderEmail`) is switched off in the script by default.

# 3. Troubleshooting

| **Problem**                                                                       | **What to do**                                                                                                                                            |
|-----------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------|
| "`earlyBirdEnabled` is 'Yes', so all early bird values are required"                | Fill in earlyBirdTicketLimit, earlyBirdDateTimeLimit (format YYYY-MM-DD HH:mm) and earlyBirdTicketPrice, or set `earlyBirdEnabled` to 'No'.                 |
| GitHub error, or "Could not get global asset image(s)" / template download failed | Wait a few minutes and run the script again. If a file is reported as not found in GitHub, check that it exists in the repository with the expected name. |
| "A folder named … already exists"                                                 | Change `showName`, or rename or remove the existing show folder.                                                                                            |
| Image links in the Counters tab are empty                                         | Check the file names and the folder they are in (sections 1.3 and 1.4), then run `setupSheets` again.                                                       |
| Warning that the Counters tab does not match the layout                           | The tab was created by an older version. Delete the Counters tab and run `setupSheets` again. Note the values in B1 and B12 first, as they reset.           |
| Emails are not being sent                                                         | Check that `createAllTriggers` was run, and check the daily email and trigger limits for your account type (section 1.1).                                   |
