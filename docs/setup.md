# Production setup (Windows)

The bot runs once a day on an ordinary Windows laptop or PC, started by Task Scheduler. It needs no server and no paid software.

## 1. Install

1. Install **Python 3.11 or newer** from python.org. During setup, tick *Add python.exe to PATH*.
2. Copy this repository to a folder, for example `C:\SalesPipeline`.
3. In that folder, open a terminal and run:

   ```bat
   pip install -r requirements.txt
   ```

## 2. Configure

1. Copy `config\config.example.json` to `config\config.json`. Git ignores `config.json`, so it never leaves the machine.
2. Edit these settings:
   - `senders`: the addresses that send the report mails. All other mails are ignored.
   - `imap.user`, `smtp.user`: the mailbox the reports arrive in.
   - `smtp.to`: who receives the health mail.
   - `groups`: one entry per report file. Each entry has the file name as it arrives (without `.xlsx`), the handler, and `required` (true or false).
   - `schedule`: the start time, cut-off time and SLA target.
3. **Mail password.** For Gmail, turn on 2-step verification and create an *app password*. Then store it in an environment variable, never in the config file:

   ```bat
   setx GMAIL_APP_PASSWORD "abcd efgh ijkl mnop"
   ```

   Open a new terminal afterwards so the variable is visible.

## 3. Reference files (`data\reference\`)

| File | Columns | Purpose |
|---|---|---|
| `Team_Mapping.xlsx` (sheet `Team`) | CA NAME, TL NAME, SM NAME, Team Type (PV/EV), Split login (PV/EV), Note | CA → TL → SM. Edit it when teams change; the history is kept |
| `CA_ID_List.xlsx` (first sheet) | CA ID, CA Name | CRM login → person. Give a shared login one row per person |
| Sales register (optional) | Chassis No, Actual CA | Who really sold each vehicle. Publish it as CSV and put the link in `ssi.actual_ca_csv_url`; the last good copy is cached |

If the register is a Google Sheet, a helper tab with the formula below keeps it in the right shape. Here the chassis number is in column N and the actual seller in column W:

```
={"Chassis No","Actual CA";FILTER({Data!N2:N,Data!W2:W},Data!N2:N<>"")}
```

**File → Share → Publish to web** (that tab, CSV) then gives you the link. Treat that link like a password: anyone who has it can read the sheet.

## 4. Test without the mailbox

1. Set `"source_mode": "folder"` in `config.json`.
2. Put one day's report files in `data\test_inbox\`.
3. Run:

   ```bat
   python src\mail_bot.py --once --no-email
   ```

   The health mail is saved as `logs\health_<date>_<time>_<STATUS>.html`. Open it in a browser.
4. Set `source_mode` back to `"imap"`.

## 5. Load the older months (once)

Put each month's complete "Bookings – Current Month" files in `data\backfill\<yyyy-mm>\`. Then run:

```bat
python src\backfill.py --dry-run     & rem shows what would happen
python src\backfill.py               & rem does it (bookings already in the database are never changed)
```

## 6. Schedule it

In Task Scheduler, choose **Create Task** (not *Basic Task*) and set:

| Tab | Setting |
|---|---|
| General | Run only when the user is logged on |
| Triggers | Daily at **08:30**, the same as `schedule.start_time` |
| Actions | Start a program: `C:\SalesPipeline\run_bot.bat`; *Start in*: `C:\SalesPipeline` |
| Conditions | **Untick** "Start the task only if the computer is on AC power" and "Start only if the following network connection is available" (the bot waits for the network itself); **tick** "Wake the computer to run this task" |
| Settings | **Tick** "Run task as soon as possible after a scheduled start is missed"; "If the task is already running: **Do not start a new instance**" |

`run_bot.bat` reopens itself minimised with the title *"Bookings bot – running – do NOT close"*. It closes itself when the day is done.

"Wake the computer to run this task" only works when wake timers are allowed. Check Control Panel → Power Options → Change plan settings → Change advanced power settings → Sleep → **Allow wake timers: Enable**. A computer that is switched off cannot wake up, so keep the laptop plugged in and asleep overnight, not shut down. When the bot does start late, the health mail says so.

Only one copy of the bot runs at a time. If you start it by hand while the scheduled one is still waiting for files, the second copy says *"Another copy of the bot is already running"* and stops. Close the bot window first if you want to run it by hand.

## 7. Every day

- Read the health mail.
  - **GREEN:** refresh Power BI.
  - **RED:** the mail says what is missing or failed, and the bot keeps retrying until 20:00.
- To recover by hand after fixing a problem:

  ```bat
  python src\mail_bot.py --once --final
  ```

- To check that every workbook can be read, run `python src\check_workbooks.py`. If one is damaged, run `python src\check_workbooks.py --restore` and then the commands it prints.
- To reload one day from scratch:

  ```bat
  python src\mail_bot.py --date 2026-10-05 --once --final --reset-state
  ```

## 8. Add a new report type

1. Copy `src\handlers\_template.py` to `src\handlers\<name>.py` and fill in `process()`.
2. Add the file to `config.json`:

   ```json
   {"name": "pv <file name>", "handler": "<name>", "required": false}
   ```

The next run picks it up. The file grid on the Data Operations page shows the new file automatically.
