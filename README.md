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
| Board.html | the "This week" page: the paper weekly grid, live (same address, add ?view=week) |
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

- **Log** tab: one row per tap with the exact time, campus, category, In person or Remote.
- **Gate** tab: one row per campus, day and time block. Entering a block again replaces it.
- One tab per campus: the paper weekly sheet, filled in live. Type any date in the yellow cell to see that week.
- **Summary** tab: this month by campus, by category, and the busiest hours. Type any date in the yellow cell to see that month.
- **Buttons** tab: the list of buttons. Change a name, the helper text, the order, or set Show to no. The page picks it up within a minute.
- **Campuses** tab: the campus names and each one's gate time blocks.

## Things worth knowing

- Undo takes back the last tap for two minutes, nothing else. After that, remove the row in the Log tab.
- If the wifi drops, taps wait on the device and send themselves when it is back. The top right corner says how many are waiting.
- The page stops working the day the Google account that deployed it is deleted. Deploy it from an account the library keeps.
- For a weekly dated copy of the Log and Gate tabs in a Drive folder called "Desk Stats backups", run the function **installWeeklyBackup** once from the editor.
- The time zone is one word in appsscript.json (America/Chicago). Change it there if the college ever moves.
