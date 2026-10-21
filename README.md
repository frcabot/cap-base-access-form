# CAP Base Access & Graduation Guest Form

A Google Form plus a Google Sheet, run by one Apps Script file (`BaseAccessForm.gs`). It collects the identity information a military installation's security office needs to vet visitors for a CAP activity held on base, such as an encampment. It also collects each attendee's graduation guest list. Then it does the follow-up work that usually takes an administrator weeks by hand.

It was built for and used at a wing winter encampment with about 475 registrants: staff, students, alternates and their families.

---

## What it does

**One form, built for parents, cadets and first-time visitors**
- It is a branching form. Each person sees only the questions that apply to them:
  - An adult attendee gives their photo ID.
  - A cadet under 18 who is not driving gives nothing.
  - A driver gives a driver's license.
  - A guest with a license without a star ("Federal Limits Apply") is asked for a Social Security number and a second ID.
  - A non-citizen is asked for their immigration document and passport number.
- **Military / DoD ID and federal PIV card holders skip all ID questions.** They usually do not need an installation access request. The form still records them, so the administrator knows they are covered and whether they have guests.
- **Guests:** up to 8 per attendee (adjustable). Guests under 18 who are not driving give only a name and date of birth. Parents with several cadets list guests on one form only. "Not sure yet" lets people submit now and add guests later.
- **Plain-language wording, written for 12-year-olds, parents who struggle with forms, and people whose first language is not English.**
  - "Driver" means the person behind the wheel. Every "driver" question is phrased to stop parents answering "yes" because they are driving their cadet.
  - The help text tells people how to spot a REAL ID star.
- **Personal links:** each attendee gets a link that already has their CAPID, rank, name, home address and age filled in.
- **Attendee check:** the CAPID question only accepts CAPIDs on the roster, so the person entered is a registered attendee.

**Automatic checking, explained in plain words**

Each submission is checked against the installation's usual rules, for example:
- The ID expires before the event ends.
- No expiration date was given.
- A license without a star is missing its Social Security number or second ID, or the "second ID" is not one they have now (for example "will get a REAL ID") or has no number.
- "ID issued by" does not fit the ID type: a license must be issued by a state, a U.S. passport by the United States, a foreign passport by a foreign country.
- The ID number is actually the CAPID.
- A non-citizen gave no immigration document.
- Someone under 16 is marked as a driver.
- The age given does not match Registration Zone.
- The date of birth does not match Registration Zone.
- The last page was not signed.
- A Step 1 or guest question was left unanswered.

**Emails that tell people exactly what to do**
- **After every submission or edit**, the person who filled out the form gets an email. The attendee is CC'd, and so is the parent for cadets under 18. The email is one of:
  - **"Received – STATUS: COMPLETE"**, or
  - **"ACTION NEEDED – STATUS: NOT COMPLETE"**, with numbered items. Each item says what they entered, what is wrong, and a **FIX:** line with exactly what to change, plus their personal edit link.
- **A second form for the same attendee** adds guests to the earlier list. It never replaces it. Corrections to the same guest replace the earlier entry.
- **Easy to read on any phone or computer.** Every email is sent formatted (paragraphs, numbered fix items, bold labels, clickable links, COMPLETE in green and NOT COMPLETE in red), with a plain-text copy for mail apps that do not show formatting.

**Automatic follow-ups**
- **Twice a week** (Mondays and Thursdays), staff and slotted students get a reminder if they:
  - have not submitted, or
  - submitted a form that still has problems.
- **Alternates are not chased.** They still get the confirmation or fix email whenever they submit.
- **Spacing:** nobody is emailed twice within 2 days.
- **Summary:** the administrator gets an email after each run.
- **Optional:** weekly reminders for registrations still missing **unit commander or parent/guardian approval in Registration Zone**. Commanders with several pending registrations get one combined email. Parents get step-by-step approval instructions.

**Reports for the administrator and installation security**
- **Roster tab:** who has submitted, whose own ID is on file, guest counts, and problems.
- **Submissions tab:** one row per attendee, newest first, with plain-language flags.
- **SECFOR Roster (masked):** everyone who will come through the gate, with only the **last 4 digits** of each SSN and ID number.
  - **Each person appears once**, even when siblings each list the same parents as guests (same name and date of birth = same person; both sponsors are shown).
  - **"ID issued by" is standardized** for security: a 2-letter state for licenses and state IDs ("California DMV" -> CA), "United States" for U.S. federal IDs, and a standard country name for foreign passports. Citizenship is standardized the same way in the full export. Anything not recognized is left exactly as typed. The checks and fix emails still use what people actually typed.
- **Full SECFOR export:** a separate, private spreadsheet with the full information, for sending to installation security.
- **Email log:** every email sent, with recipients and status.
- **Help Requests:** people can pick "I need help" on the form. They get an automatic answer to the common problems, and the request is emailed to the administrator.

---

## What you need

- A Google account. A wing Google Workspace account is best: it can send about 1,500 email recipients a day, compared with about 100 a day for a personal Gmail account.
- The **host installation's visitor access request form**. Use it to check the legal text and the list of acceptable IDs. Also ask its security office:
  - whether military/PIV holders are exempt,
  - whether a Social Security number is required,
  - what format they want the list in.
- *(Optional)* The event's **Registration Zone exports** (.xlsx), saved in one Google Drive folder. With them, the roster fills itself and the approval reminders work. Without them, you type or paste the roster yourself.

---

## Set up (about 30 minutes)

1. **Create a Google Sheet** in the Drive folder where you want the form to live. Name it something like `Encampment_Base_Access`.
2. **Extensions > Apps Script.** Delete the sample code and paste all of `BaseAccessForm.gs`. Click **Save**.
3. **Fill in the SETTINGS** at the top of the script:
   - the event name and dates,
   - the deadlines,
   - the installation's common and official names,
   - your name, title and email,
   - *(optional)* the Reg Zone folder ID and file names.

   `EVENT_END_ISO` must be written as `YYYY-MM-DD`. Every ID must still be valid on that day.
4. **Replace the installation legal text:** `INSTALLATION_PRIVACY_ACT`, `INSTALLATION_RULES` and `CONSENTS`, just below the settings. The included text is an example from one Army installation; use the wording on *your* installation's access request form. Also check `ALL_TYPES` (the acceptable IDs) against that form. Then set `LEGAL_TEXT_REVIEWED: true`.
5. **Run `setup`.** Pick `setup` in the function list at the top of the editor, then click Run. Approve the permissions Google asks for.
   - Building the form takes longer than Google's 6-minute limit, so setup works in rounds. When it says "NOT FINISHED YET", click Run again.
   - It usually needs 2–4 rounds. It picks up where it stopped each time.
   - Setup will refuse to start if any setting still has its example value.
6. **Reload the spreadsheet.** A **Base Access** menu appears.
7. **Fill the roster.** Use either:
   - **Base Access > Pull new people from Reg Zone exports** (and **Turn on daily Reg Zone pull** to keep it current), or
   - the Roster tab by hand, with the columns below.

   Then run **Base Access > Update CAPID list from Roster tab** and **Regenerate personal links**.
8. **Choose the sending account.** If the script will send email from a different account than the one you set it up with, sign in as that account and run **Base Access > Run emails and automation from MY Google account**. That account must be an editor of the form.
9. **Test it yourself before sending links.** Open the form from your own personal link (Roster tab) and try several paths:
   - an adult with a REAL ID,
   - a cadet under 18 who is not driving,
   - a driver,
   - a military ID holder,
   - two guests.

   Check the emails you receive, then delete your test responses in the Form (Responses > the three dots).
10. **Send the links:** **Base Access > Email personal links**. Then turn on **Turn on automatic follow-ups**.

### Roster tab columns (if you fill it in by hand)

| Column | What to put |
|---|---|
| `CAPID` | The attendee's CAPID |
| `Rank`, `First name`, `Last name` | As in eServices, for example `C/Amn`, `Jane`, `Smith` |
| `Event` | The attendee's group, for example `Staff`, `Encampment`, `ATS` |
| `Member type` | `Cadet` or `Senior` |
| `Birth month/year (Reg Zone)` | `MM/YY` (optional; used to catch a wrong date of birth) |
| `Age at event start` | A number |
| `Must submit own ID (18+ at start)` | `Yes` or `No` |
| `Email`, `Parent email` | Where confirmation and reminder emails go |
| `Home address (Reg Zone)` | Pre-fills the form (optional) |
| `Slot status` | `Staff`, `Slotted` or `Alternate`. Follow-ups chase `Staff` and `Slotted` only. |
| `Base access override` | *(optional)* Fill in for people who do NOT need the form, for example `ID on file from the fall activity` or `Not needed: retired military (DoD ID)` |

---

## Day to day

| Menu item | What it does |
|---|---|
| Rebuild tabs now | Refreshes every tab. It also runs automatically after each submission. |
| Pull new people from Reg Zone exports | Adds new registrants, updates ranks and names, and marks people who dropped out |
| Check automation status | Shows which Google account runs the emails and scheduled jobs, which jobs are on for that account (daily pull, confirmation emails, follow-ups, email queue, weekly approval reminders), when the last Reg Zone pull finished, and the time zone used for "6 AM" |
| Email personal links | The first email with each person's link (only to people not emailed yet) |
| Send reminders | Reminds staff and slotted students who have not submitted |
| Email fix requests | Re-sends the "ACTION NEEDED" email to forms that still have problems (each version once) |
| Preview / Send / Turn on automatic follow-ups | The twice-weekly reminders described above |
| Create full SECFOR export | A private spreadsheet with full SSNs, dates of birth and ID numbers for installation security. Delete it after you send it. |
| Apply wording changes | After you edit any help text or email wording in the script: pushes the new wording onto the existing form |
| Delete ALL form data | After the event: deletes every response |
| Registration reminders menu | Preview, send, or turn on weekly unit/parent approval reminders (needs the Reg Zone exports) |

---

## Customizing

- **Wording:** all text people see is in the script, mostly near the top (`HELP`, `COVER_HELP`, `DRIVER_HELP_*`, `formDescription_`, `guestCountHelp_`, `CONFIRM_MSG`, and the email functions). After editing, run **Apply wording changes**.
- **Acceptable IDs:** `ALL_TYPES` and `DRIVER_TYPES`. Edit these **before** running setup. Changing answer choices on a live form changes how older responses are read.
- **SSN rule:** `SSN_REQUIRED_FOR_ALL`.
- **Military / PIV skip:** keep it only if the host installation confirms those card holders need no access request.
- **Number of guest pages:** `MAX_GUESTS`. Set it before setup. More pages make setup take longer.
- **Follow-up timing:** `CHASE` (days between reminders, hour of the day). The run days are set in `turnOnAutoFollowUps`.

---

## Privacy and data handling

This form collects **Social Security numbers, dates of birth and ID numbers**. That is Controlled Unclassified Information (CUI) and PII. Handle it under your wing's and the installation's rules, and check with your wing before you use the form.

- **Full answers stay inside the Google Form.** They are never copied into the spreadsheet. The spreadsheet tabs show only the last 4 digits and ages.
- **The full export** is a separate spreadsheet that only you can open. Send it to installation security the way they ask, then delete it.
- **Edit links contain the full answers.** Anyone with a person's edit link can see what was entered. The script sends each link only to the person who filled out the form, the attendee and the parent. Do not forward these emails or add more recipients.
- **Limit who can open the form and the spreadsheet** to the people who administer it.
- **After the event,** run **Delete ALL form data** and delete any exports.

---

## Limits and tips

- **Daily email limit:** each Google account has one. When it is reached, emails wait on an "Email queue" tab and go out automatically within the next hours. `EMAIL_RESERVE` keeps room for confirmations.
- **Scheduled jobs belong to one Google account.** Turn on the daily pull and other automatic jobs while signed in as the account shown by **Check automation status**. If a daily pull was set up from a different account, it turns itself off and emails you once instead of silently doing nothing.
- **Change the form at quiet times.** Changing questions or answer choices while people are filling it out can make some of their answers disappear. "Apply wording changes" only changes text, but run it late at night or early in the morning anyway.
- **People with two or more forms** for the same attendee are listed for you to check by hand instead of being emailed automatically. One edit link might be the wrong one.
- **The checks catch common mistakes, not everything.** Installation security makes the final access decision.
