# Desk Stats

A tap counter for the front desks of Motlow State's four campus libraries
(Smyrna, Moore County, Fayetteville, McMinnville). Staff tap one button per
question; gate counts are typed in per time block. Everything lands in one
Google Sheet, which also carries a live copy of the paper weekly sheet for
each campus and a month summary. A second page shows a live board.

It is a Google Apps Script project: no server, nothing to install, runs on
the college's Google accounts.

## The four files

| File | What it is |
|---|---|
| Code.gs | the brain: builds the sheet, serves the pages, saves every tap |
| Index.html | the tap page staff use on the iPad or phone |
| Board.html | the "This week" page: the paper weekly grid, live, for this week or any earlier one (same address, add ?view=week) |
| appsscript.json | two settings: the time zone and how the page is shared |

## Putting it into Google (about five minutes)

1. Go to script.google.com and click **New project**.
2. Click the name **Untitled project** at the top left, type **Desk Stats**, press Enter.
3. In the editor you will see a file called **Code.gs** with a few lines in it. Click into it, select everything (Ctrl+A), delete it, and paste the whole contents of this folder's **Code.gs**. Press Ctrl+S.
4. On the left, next to **Files**, click **+** then **HTML**. Type **Index** and press Enter. Select everything in the new file, delete it, paste the whole contents of **Index.html**. Ctrl+S.
5. Do the same for **Board**: **+**, **HTML**, name it **Board**, paste **Board.html**. Ctrl+S.
6. Click the gear icon on the left (**Project Settings**). Tick **Show "appsscript.json" manifest file in editor**. Click **< >** (Editor) on the left again. Open **appsscript.json**, replace its contents with this folder's **appsscript.json**. Ctrl+S.
7. Click **Code.gs** in the Files list so it is the open file (the function list only shows the functions of the file that is open). At the top of the editor there is a dropdown next to **Run** and **Debug**; it now lists **setup**, **doGet** and the rest. Pick **setup**, then click **Run**. The first time, Google asks you to approve the script: click **Review permissions**, pick your account, click **Advanced**, then **Go to Desk Stats (unsafe)**, then **Allow**. It says unsafe only because the script is yours and not published in Google's store.
8. When it finishes, the **Execution log** at the bottom shows the address of the new spreadsheet, named **Desk Stats**. It is in your Google Drive.
9. Click **Deploy** (top right) then **New deployment**. Click the gear next to **Select type** and choose **Web app**. Description: Desk Stats. **Execute as: Me**. **Who has access: Anyone with Google account** (for the college install pick "Anyone within" the college instead). Click **Deploy**. Copy the **Web app URL**.
10. Open that address on the iPad. Tap **Share** then **Add to Home Screen**. The icon opens on the campus choice screen.

## While the code is still changing: use the test address

A deployment is a snapshot. After you paste new code, the address ending in
**/exec** keeps showing the old snapshot until you publish again. Two ways
round that:

- **Test address (no publishing needed):** Deploy > **Test deployments**.
  Copy the address ending in **/dev**. It always runs the code you last
  saved (Ctrl+S). Use this while we are fixing things. Only you can open it.
- **Publish the new code to the real address:** Deploy > **Manage
  deployments** > the pencil icon > under **Version** choose **New
  version** > **Deploy**. The /exec address now serves the new code.

## The This week page

Same address as the tap page with **?view=week** on the end (the This week
button on the tap page opens it). It reads every recorded day once, then
opens on the current week. The week label above each campus grid is a
button: it opens a small window where the arrows step back one week at a
time, as far back as the first week with anything recorded, and **Show**
switches at once, with no waiting. A week with nothing recorded still shows,
as zeros, and the window says so. The page re-reads the sheet every minute;
the top right corner always says what is going on: Loading the sheet,
Refreshing, Saving, Up to date, Could not read the sheet, or The last
change was not saved.

While the page reads the sheet for the first time, and whenever **Refresh
now** is pressed, a small Loading window sits in the middle of the page
until the numbers are in; the once-a-minute re-read only says Refreshing in
the corner and never redraws the grid under your finger.

**Edit numbers** (top right) works once a single campus is chosen (it is
greyed out on All campuses, so a number is always changed on one named
campus). Edit mode looks like the tap page: the grid's header turns green,
a yellow strip names the campus being edited, and every question number
becomes a block with a down arrow, the number, and an up arrow, green for
In person and yellow for Remote. The up arrow adds one and the down arrow
takes one off; presses on the same number within about a second go to the
sheet as one change, the block is striped until the sheet confirms, and
every total follows. A gate count becomes a raised button that opens the
keypad, and saving replaces that block's number (the old one stays in the
Gate history tab). Nothing already in the sheet is erased: a question change
is one correction row in the Log tab (Count = the change, up or down,
Source = edit, who made it, no clock hour). If the sheet refuses a change, a
red box says why and the grid keeps the sheet's number. Days that have not
happened yet stay plain and cannot be changed.

## The Trends page

Same address as the tap page with **?view=trends** on the end (the Trends button on the tap page opens it). It reads
the whole sheet once, then every filter (campus, term, last 90 days) works
instantly. It shows questions per week, visitors per week, questions by
month and category, campus comparison, the weekday pattern, visitors
against questions day by day (with how strongly they move together), the
hour-of-day pattern (from real taps only), category mix per campus, the
remote share by month, and gate counts by time block, plus a "what stands
out" list written in plain sentences.

## Loading the old LimeSurvey weeks (one time)

The folder holds **history-log.csv** and **history-gate.csv**, made from the
LimeSurvey export by import/limesurvey_to_history.py, and
**history-report.txt** listing every row it dropped or changed.

1. Open drive.google.com and drag **history-log.csv** and **history-gate.csv** into it (anywhere; the script finds them by name).
2. In the script editor click **Code.gs**, pick **importHistory** in the function dropdown, click **Run**. Approve the Drive permission if asked.
3. The **Execution log** says how many rows went into Log and Gate. Running it again adds nothing twice.
4. Open the dashboard and the campus weekly tabs: the old weeks are there, marked "LimeSurvey" in the Source column.

## Where the numbers go

- **Log** tab: one row per tap with the exact time, campus, category, In person or Remote, and one row per correction made on the This week page, holding the difference (Source: edit).
- **Gate** tab: one row per campus, day and time block, holding the current number. Entering a block again replaces it here, so every total stays right.
- **Gate history** tab: every gate entry ever made from the page, in order, with the time, who, the number, and the number it replaced. Nothing here is ever overwritten, so a wrong entry and its correction are both on record. Totals never read this tab. A workbook from before this round gets the tab by itself on the first gate entry (or run **setup** again).
- One tab per campus: the paper weekly sheet, filled in live. Type any date in the yellow cell to see that week.
- **Summary** tab: this month by campus, by category, and the busiest hours. Type any date in the yellow cell to see that month.
- **Buttons** tab: the list of buttons. Change a name, the helper text, the order, or set Show to no. The page picks it up within a minute.
- **Campuses** tab: the campus names and each one's gate time blocks.

## How counting works (the part that matters for every total)

- Each green or yellow block reads "down arrow, count, up arrow". Up adds one. Down takes one off.
- A take-off is not an erase: it is its own row in the Log tab with Count -1, the time, and who did it. Every total anywhere adds up the Count column, so a -1 row lowers the day's number and the record of it stays.
- The sheet refuses a take-off that would push that day's count for that campus, kind, and mode below zero, whoever taps and from whichever device. On the page the down arrow goes grey at zero.
- Each arrow ignores a second tap for one and a half seconds (a grey shade sweeps across it). Only that arrow: a different block, or the other arrow on the same block, still works.
- "Last 10 taps" on the tap page lists the newest taps at that campus with time, who, kind, mode and +1 or -1.
- Never count rows in this sheet. A row is +1, -1, a typed-in history day's number, or a change from the This week page (the difference, up or down). Sum the Count column.

## If numbers were saved by an earlier round (one-time repair)

Rows written by the versions before DS-2026-09-11-09 could be filed under
the wrong day, and a gate block could be saved twice. In the script editor
click **Code.gs**, pick **repairDates** in the function list, click **Run**.
The Execution log says how many rows it re-dated and how many duplicate
gate rows it removed. Running it again changes nothing more.

## How saving works (so the corner makes sense)

- A tap is written to the device's storage and sent to the sheet at once. Taps made in a burst go up in one call, in order, and each row carries the moment of its own tap (the device's clock, aligned to Google's when the page opened), not the moment the batch reached the sheet.
- The corner says **Saving N** while a send is in progress, **Connected** when nothing is waiting, and **Cannot reach the sheet** only when a send has actually failed; the red box then shows the sheet's own words, and the page keeps trying by itself.
- Every minute, and whenever the tab comes back into view, the page re-reads today's counts from the sheet. The sheet is the truth; the page adds only the taps the sheet has not received yet.
- Each browser tab keeps its own queue. A tab closed with taps still waiting leaves them for the next open tab to send.

## Things worth knowing

- If the wifi drops, taps wait on the device and send themselves when it is back. The top right corner says how many are waiting.
- The page stops working the day the Google account that deployed it is deleted. Deploy it from an account the library keeps.
- For a weekly dated copy of the Log, Gate and Gate history tabs in a Drive folder called "Desk Stats backups", run the function **installWeeklyBackup** once from the editor.
- The time zone is one word in appsscript.json (America/Chicago). Change it there if the college ever moves.
