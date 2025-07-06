# PavOra

**PavOra** is a Google Apps Script-based web app embedded in a Google Sheet that allows you to insert, edit, delete, and view reservations for various facilities. It’s been designed especially for managing events in fairgrounds, congress centers, and multi-room venues.

> Released under the MIT License

---

## Key Features

- Easy setup of structures, event types, and time slots via a dedicated configuration sheet
- Automatic conflict detection: check if a structure is already occupied, and by whom
- Multi-language support (currently English and Italian)
- Multi-user access and permission management
- Two management modes: **Event Management** and **Daily Room Management**
- Multiple views: 
  - Gantt-like calendar
  - Daily summary
  - Daily layout/plan
  - Editable reports
- Built-in backup
- Email reminders for expired pending or offer events
- Export to Excel in one click
- And much more!

---

## Step-by-Step Installation Guide

### 1. Gmail Account and Calendar

- Use your or create a new Gmail account (in the second case you have to use a separate browser with only the new profile for this).
- Create a Google Calendar or use your default one.
- Copy the calendar **ID** (./images/a_copy_your_calendar_ID.jpg).

---

### 2. Copy **PavOra Custom Settings** Configuration Sheet

- Open [PavOra Custom Settings Sheet](https://docs.google.com/spreadsheets/d/18j_d2ApLsIHOnTBbThxKV3u61VtOKzndCXT6Vlb36bw/edit?gid=0).
- Make a copy (./images/a_make_a_copy_of_PavoraCustomSettings.jpg).
- Confirm the copy (./images/b_change_the_name_and_make_a_copy.jpg)
- Save the Sheet ID from the URL (./images/c_copy_the_ID_of_the new_sheet.jpg).

---

### 3. Copy **PavOra** Main Sheet

- Open [PavOra Sheet](https://docs.google.com/spreadsheets/d/1kZ1ZfmN9Vyy3ZNuEFUG4bWITTngTa8jzg1YNJyGyfGI/edit?usp=sharing).
- Make a copy and save its ID (like step 2).

---

### 4. Configure Settings Sheet

- Open your copy of the Custom Settings Sheet.
- In the **DataSettings** sheet, replace red fields with your Calendar ID, Sheet IDs, and user emails.
- Choose the language (`en` or `it`).

---

### 5. Configure Script in the Main Sheet

#### a. Open Script Editor
- Go to `Extensions > Apps Script` (./images/d_enter_in_GoogleAppsScript_mode.jpg).

#### b. Set Global Variables
- In `01_main.gs`, update `IDPavoraCustomSettings` with your custom settings sheet ID (./images/e_change_global_variables__in_GoogleAppsScript_code.jpg).
- Optionally, configure `IDAliasEmail` to anonymize Gmail addresses, like the original one.
- Save with `CTRL+S`.

#### c. Add Triggers
- In the **Triggers** tab, add the function in the image (./images/f_add_triggers__in_GoogleAppsScript_code.jpg):
  - `userWriteReadCalendar` on open → avoids unwanted edits from users in the calendar
  - Optionally `checkEventsAndSendEmails()` (time-based: daily) → sends email reminders
- Authorize the script when prompted
  - Choose your Gmail account (./images/g_accept_permission_to_execute_app.jpg).
  - Click on `Go to PavoraScripts (unsafe)` (./images/h_go_to_pavoraScript.jpg).
  - Continue with `Sign in to PavoraScripts` (./images/i_go_on_with_continue.jpg).
  - Finally allow it (./images/l_allow_it.jpg).

#### d. Reload Sheet
- Reload PavOra Sheet. The `PavOra Menu` should now appear (./images/m_reload_and_see_menu_it.jpg/).

#### e. Grant Permissions
- Add your Gmail to the online sheet ((./images/m_add_administrator_in_the_online_sheet.jpg)).
- Run `Manage User Permission (only Admin)` from the main menu (./images/m_reload_and_see_menu_it.jpg/).

---

### 6. Final User Setup

- Users must accept the shared calendar in Gmail (./images/n_add_calendar..jpg).
- Then open PavOra Sheet, run a menu action (e.g., Reload Sidebar), and authorize the script:
  - Choose your Gmail account (./images/g_accept_permission_to_execute_app.jpg).
  - Click on `Go to PavoraScripts (unsafe)` (./images/h_go_to_pavoraScript.jpg).
  - Continue with `Sign in to PavoraScripts` (./images/i_go_on_with_continue.jpg).
  - Finally allow it (./images/l_allow_it.jpg).

---

## Repository Contents

- All code exported with `CLASP`, keeping Google Apps Script project structure.
- Includes:
  - `01_main.gs`
  -
  -
  -

---

## 📎 Related Files (Shared in View-Only Mode)

- PavOra Custom Settings Sheet – [Copy & Customize](https://docs.google.com/spreadsheets/d/18j_d2ApLsIHOnTBbThxKV3u61VtOKzndCXT6Vlb36bw/edit)
- PavOra Sheet – [Copy & Use](https://docs.google.com/spreadsheets/d/1kZ1ZfmN9Vyy3ZNuEFUG4bWITTngTa8jzg1YNJyGyfGI/edit)

---

## 🌐 Multilingual Support

- UI is currently available in **English** (90%) and **Italian** (100%)
- Easily extendable to other languages by adding new entries in `06_multiLanguageSetup.gs`

---

## 🔗 Links

- GitHub: [github.com/tuo-username/pavora](https://github.com/tuo-username/pavora)
- Website: [coming soon]

---

## 🙌 License

[MIT License](LICENSE)

---

_You're welcome to fork, adapt and improve PavOra. If you do, let me know!_
