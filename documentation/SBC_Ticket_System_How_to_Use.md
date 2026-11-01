# SBC Ticket System

## How to Use

This guide explains how to run a ticket reservation once the system has been set up (see the [**Setup Guide**](SBC_Ticket_System_Setup_Guide.md)). It is written for the **show admin**, the person who manages reservations and payments. The people buying tickets are called **showgoers**.

The show admin works mainly in the **Tracking** tab of the response spreadsheet. Everything else – reservation IDs, prices, emails, tickets – is done by the system.

# 1. The reservation process

The process has eight steps. The system runs on timers, so each automatic step happens within a few minutes to an hour (see section 5).

| **\#** | **Who**    | **What happens**                                                                                                              |
|--------|------------|-------------------------------------------------------------------------------------------------------------------------------|
| 1      | Show admin | Copies the **responder link** of the form and shares it with showgoers.                                                       |
| 2      | Showgoer   | Fills in the form to reserve tickets.                                                                                         |
| 3      | System     | Processes the entry and sends a **Reservation Received** email with the payment instructions.                                 |
| 4      | Showgoer   | Pays and **replies to the Reservation Received email** with the proof of payment.                                             |
| 5      | Show admin | Checks the payment and sets the status in the Tracking tab to **Payment Confirmed**, using the reservation ID from the email. |
| 6      | System     | Sends a **Confirmation** email with the QR-PDF tickets and lists the tickets in the **Tickets** tab.                          |
| 7      | Show admin | If the reservation is not paid within the allowed period, sets the status to **Cancelled**.                                   |
| 8      | System     | Sends a **Cancellation** email to the address in the cancelled entry.                                                         |

# 2. Step by step

## Step 1 – Share the responder link

1.  Open the form for editing. The link is in the **Form Edit URL** row of the **References** tab, and the form is also in the show folder in Google Drive.

2.  At the top of the page, click the **Published** button.

3.  Click **Copy responder link**.

4.  Share the link with your audience (social media, messages, posters, and so on).

The form is open to anyone with the link, so showgoers do not need a Google account.

## Step 2 – Showgoers reserve tickets

The showgoer fills in the form: name, email address, mobile number, city, number of tickets (1 to 4), payment option, and consent. Each submission appears as a new row in the **Tracking** tab.

## Step 3 – Reservation Received email

The system gives each new row a **Reservation ID**, works out the **Total Price** and sets the **Status**:

- **On Hold** – the tickets are reserved and the showgoer is sent the **Reservation Received** email.

- **Waitlisted** – the show is full (see section 3). The showgoer is sent a waitlist email instead.

The Reservation Received email contains the payment instructions for the chosen payment option (bank transfer, Maya, GCash or PayPal), the QR code and the payment deadline. The **Reservation Received Email** column of the Tracking tab shows **SENT** once it has gone out.

## Step 4 – Proof of payment

The showgoer pays and replies to the Reservation Received email, attaching the proof of payment. The reply arrives in the mailbox the emails are sent from. The reservation ID is in the subject line of the email.

## Step 5 – Confirm the payment

1.  Open the reply and check the proof of payment against your bank, Maya, GCash or PayPal records.

2.  In the **Tracking** tab, find the row with the same **Reservation ID** as in the email subject. Check the total price as well.

3.  In the **Status** column, choose **Payment Confirmed** from the dropdown.

## Step 6 – Confirmation email and tickets

Within the hour, the system generates one QR-PDF ticket per reserved ticket and emails them to the showgoer in the **Ticket Purchase Confirmed** email. Each ticket is also listed in the **Tickets** tab (name, email, reservation ID, ticket number and ticket ID). The **Confirmed Email** column of the Tracking tab shows **SENT**.

## Steps 7 and 8 – Cancel unpaid reservations

1.  The Reservation Received email states the payment deadline. The system does **not** cancel unpaid reservations by itself.

2.  After the deadline, check the Tracking tab for entries that are still **On Hold** and unpaid.

3.  For each of them, choose **Cancelled** in the **Status** column.

The system then sends the **Cancellation** email to the address in the entry, and the **Cancelled Email** column shows **SENT**. Cancelling also frees the seats (see section 3).

> **Tip:** Change only the **Status** column in the Tracking tab. Columns A to H come from the form, and the other columns are filled by the system. Do not delete or rearrange rows.

## Status reference

| **Status**              | **Meaning**                                                                                                                                      |
|-------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| On Hold                 | Tickets are reserved and waiting for payment. Set by the system.                                                                                 |
| Payment Confirmed       | Payment received. **Set by the show admin.** Triggers the confirmation email and the tickets.                                                    |
| Cancelled               | Reservation cancelled. **Set by the show admin.** Triggers the cancellation email.                                                               |
| Waitlisted              | The show is full; the showgoer is on the waiting list. Set by the system.                                                                        |
| On Hold - From Waitlist | A waitlisted entry that has moved up because seats became free. The showgoer is sent a "Reservation Opened" email with the payment instructions. |
| Ticket/s Claimed        | The tickets have been claimed at the show. Counts the same as Payment Confirmed.                                                                 |
| Invalid                 | The entry or ticket is not valid. It is not counted as a seat. The system also uses it for a ticket that could not be generated correctly.       |

# 3. Ticket limit and waitlist

The system compares the number of tickets already taken against the **Ticket Limit** every time a new reservation comes in. The limit is set in the **Counters** tab (field **Ticket Limit**).

### What counts as a seat

Tickets count as taken when the entry is **On Hold**, **On Hold - From Waitlist**, **Payment Confirmed** or **Ticket/s Claimed**. A cancelled reservation keeps its seats until its **Cancellation email has been sent**. The Counters tab shows the running totals, including **Ticket Slots Remaining**.

### When the show is full

- While seats are taken below the limit, a new reservation is set to **On Hold**.

- Once the limit is reached, all following reservations are set to **Waitlisted** and the showgoer receives the **Waitlisted** email. No payment is requested at this point.

### How waitlisted entries move up

Waitlisted entries are handled **first come, first served**, in the order they appear in the Tracking tab. They are served before any new reservation. When seats are free, the next waitlisted entries are changed to **On Hold - From Waitlist** and the showgoer receives the **Reservation Opened** email with the payment instructions. Seats become free in two ways:

- **Cancellations.** When the show admin cancels reservations and the total goes below the ticket limit, waitlisted entries move up as soon as the cancellation emails have been sent.

- **A higher Ticket Limit.** The show admin can increase the **Ticket Limit** in the Counters tab, for example when the show moves to a bigger venue. The extra seats are given to the waitlist in the same way.

Entries that move up have the same payment period and cancellation rules as other reservations.

> **Good to know:** The system only checks whether seats are still available **before** adding a reservation. A reservation of several tickets that starts just below the limit is accepted in full, so the total can go slightly over the limit (for example, with one seat left, a reservation of 4 tickets is accepted).

# 4. Early bird pricing

Early bird pricing offers a lower ticket price for a limited time or a limited number of tickets. It is turned on in the setup inputs (see the Setup Guide) and shown in the Counters tab:

| **Counters field**                    | **Meaning**                                                                                                             |
|---------------------------------------|-------------------------------------------------------------------------------------------------------------------------|
| **Early Bird Enabled**                | Yes while early bird pricing is active. No if it is not used. The system changes it to `Exceeded` when a limit is passed. |
| **Early Bird Ticket Number Limit**    | How many tickets can be sold at the early bird price.                                                                   |
| **Early Bird Ticket Date/Time Limit** | The last moment the early bird price applies.                                                                           |
| **Early Bird Ticket Price**           | The early bird price per ticket.                                                                                        |

### How it works

- While **Early Bird Enabled** is Yes, each new reservation is priced at the **early bird price**.

- The early bird price stops as soon as **either** limit is passed: the number of early bird tickets sold reaches the **ticket number limit**, **or** the reservation was submitted after the **date/time limit**.

- When that happens, **Early Bird Enabled** changes to `Exceeded`, and all later reservations use the **normal ticket price**. It does not switch back by itself.

- The date and time check uses the time the showgoer **submitted the form**, not the time the system processed it. A reservation submitted before the cutoff keeps the early bird price even if it is processed later.

- Early bird reservations are marked **Yes** in the **Early Bird?** column of the Tracking tab. The number of early bird tickets sold is kept in **Early Bird Tickets Sold** in the Counters tab.

- A reservation keeps its price. A waitlisted entry that moves up later keeps the price it was given when it was received.

- For PayPal payments, the PayPal transaction fee per ticket is added to the price.

> **Good to know:** The ticket number limit is checked **before** a reservation is priced. A reservation that starts just below the limit gets the early bird price for all its tickets, so slightly more tickets than the limit can be sold at that price. Cancelled early bird tickets are not given back to the early bird count.

# 5. When things happen

The system runs on timers, so emails are not instant. Allow for the times below.

| **What**                                                  | **Runs**         | **Expect**                                        |
|-----------------------------------------------------------|------------------|---------------------------------------------------|
| New entries processed (ID, price, status, waitlist moves) | Every 15 minutes | Within about 15 minutes of the form submission    |
| Reservation Received / Reservation Opened email           | Every 30 minutes | Up to about 45 minutes after the form submission  |
| Waitlisted email                                          | Every 5 minutes  | Within about 20 minutes of the form submission    |
| Confirmation email with tickets                           | Every hour       | Within about an hour of marking Payment Confirmed |
| Cancellation email                                        | Every hour       | Within about an hour of marking Cancelled         |

# 6. If something does not work

| **Problem**                                             | **What to check**                                                                                                                                                 |
|---------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| A new entry has no Reservation ID                       | Wait for the next 15-minute run. If it stays empty, check that `createAllTriggers` was run (see the Setup Guide).                                                   |
| The showgoer did not get the Reservation Received email | Check the **Reservation Received Email** column. If it is empty after an hour, check the email address in the entry and the daily email limit of the account.     |
| No confirmation email after Payment Confirmed           | The Reservation Received email must have been sent first. Wait up to an hour, then check the **Confirmed Email** column and the **Tickets** tab.                  |
| A waitlisted entry does not move up                     | Seats must be free (see **Ticket Slots Remaining** in the Counters tab). Cancelled reservations only free their seats after the Cancellation email has been sent. |
| The showgoer cannot open the form                       | Check that the form is published and that the responder link was copied from the **Published** menu.                                                              |
