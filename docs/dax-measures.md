# DAX measures

The report has more than 150 measures. Each measure lives in the measure table of its subject, inside a display folder, and has its description filled in. This page shows the formulas exactly as they are written in the report (one line each), followed by the catalogue of the measures from the first build phase.

## Follow-up lists (calculated columns in Open_Now)

```dax
Follow_Up_Order = VAR enq = Open_Now[Form_Order] IN { 1, 2 } VAR recent = Open_Now[Age_Days] <= 30 VAR lc = Open_Now[Lead Classification] RETURN SWITCH ( TRUE (), Open_Now[Is_Long_Wait] = 1, 1, Open_Now[Is_Stale_Enquiry] = 1, 2, enq && recent && lc IN { "Hot", "Warm" } && Open_Now[Testdrive] <> "Given", 3, enq && recent && ( lc = "Unspecified" || ISBLANK ( lc ) ), 4, 5 )
Follow_Up = SWITCH ( Open_Now[Follow_Up_Order], 1, "Booked, waiting over 30 days", 2, "Enquiry open over 30 days", 3, "Hot or Warm, no test drive (last 30 days)", 4, "Lead class not set (last 30 days)", "On track" )
```

`Follow_Up` is sorted by `Follow_Up_Order`. One list is picked with a slicer, and the follow-up page shows only that list, sorted by days open.

## Customer experience (SSI)

```dax
SSI Avg Rating = CALCULATE ( AVERAGE ( SSI_Answers[Rating] ), SSI_Answers[Q_Type] = "Rating 0-10", SSI_Surveys[Is_Done] = 1 )
SSI Points Lost = CALCULATE ( - SUM ( SSI_Answers[Points] ), SSI_Answers[Has_Penalty] = 1, SSI_Surveys[Is_Done] = 1 )
SSI Times Penalised = CALCULATE ( DISTINCTCOUNT ( SSI_Answers[CSR#] ), SSI_Answers[Has_Penalty] = 1, SSI_Surveys[Is_Done] = 1 )
SSI Customers To Call Back = CALCULATE ( COUNTROWS ( SSI_Surveys ), SSI_Surveys[Is_Done] = 1, SSI_Surveys[Total Score] < 930 ) + 0
Days Sale To Delivery Avg = CALCULATE ( AVERAGE ( SSI_Surveys[Days_Sale_To_Delivery] ), SSI_Surveys[Is_Done] = 1 )
SSI Period = VAR a = MIN ( SSI_Surveys[Survey_Month] ) VAR b = MAX ( SSI_Surveys[Survey_Month] ) RETURN IF ( a = b, FORMAT ( a, "mmm yyyy" ), FORMAT ( a, IF ( YEAR ( a ) = YEAR ( b ), "mmm", "mmm yyyy" ) ) & " – " & FORMAT ( b, "mmm yyyy" ) )
```

Other SSI measures, described by definition:

- **SSI Score:** the average Total Score of completed surveys.
- **Survey counts:** SSI Surveys / Surveys Done.
- **SSI Score Last Month.**
- **Target measures:** SSI Target (980) and SSI Target Met % (the share of surveys at 980 or more).
- **Box measures:** SSI Top Box % and SSI Bottom Box Surveys.
- **Penalty measures:** Penalty per Survey (points lost ÷ surveys), Penalised Surveys and No Answers.
- **Days measures:** Days Delivery To Call and Days Sale To Delivery (both medians).
- **Colour measures:** SSI Score Font and SSI Score Fill (orange below 930, green at 980 and above).

### Latest month vs year to date (Overview card)

On the Overview page, the SSI card shows the latest survey month with the year-to-date figure underneath. These numbers ignore the Month slicer, so the card always answers "how did we do most recently?". Year to date starts on 1 April of the financial year of the latest survey month. When the data holds only this year's SSI report, that is every row. When last year's final file is loaded too, as in the `--full` demo, the filter keeps last year out.

```dax
SSI Latest Month = CALCULATE ( MAX ( SSI_Surveys[Survey_Month] ), REMOVEFILTERS ( 'Date' ) )
SSI Score Latest Month = VAR m = [SSI Latest Month] RETURN CALCULATE ( [SSI Score], REMOVEFILTERS ( 'Date' ), SSI_Surveys[Survey_Month] = m )
SSI Surveys Latest Month = VAR m = [SSI Latest Month] RETURN CALCULATE ( COUNTROWS ( SSI_Surveys ), REMOVEFILTERS ( 'Date' ), SSI_Surveys[Survey_Month] = m, SSI_Surveys[Is_Done] = 1 )
SSI Latest Month Label = FORMAT ( [SSI Latest Month], "mmm yyyy" )
SSI Latest Month Text = FORMAT ( [SSI Latest Month], "mmm yyyy" ) & " (" & [SSI Surveys Latest Month] & " surveys)"
SSI Score YTD = VAR m = [SSI Latest Month] VAR fy = DATE ( YEAR ( m ) - IF ( MONTH ( m ) < 4, 1, 0 ), 4, 1 ) RETURN CALCULATE ( [SSI Score], REMOVEFILTERS ( 'Date' ), SSI_Surveys[Survey_Month] >= fy )
SSI Latest Font = VAR s = [SSI Score Latest Month] RETURN IF ( s < 930, "#E46C0A", IF ( s >= 980, "#2E7D32", "#252423" ) )
```

## Team performance

```dax
Team Activity = COALESCE ( [Bookings Month To Date], 0 ) + COALESCE ( [Bookings Same Days Last Month], 0 ) + COALESCE ( CALCULATE ( COUNTROWS ( Open_Now ), REMOVEFILTERS ( 'Date' ) ), 0 )
```

`Team Activity` is used as a visual filter (`is greater than 0`). It hides people who have no activity, without hiding a CA who sold nothing this month but did sell last month.

## Data operations

```dax
Ops Last Status = VAR d = MAX ( Daily_Runs[Run_Date] ) RETURN CALCULATE ( SELECTEDVALUE ( Daily_Runs[Status] ), Daily_Runs[Run_Date] = d )
Ops Status Color = SWITCH ( [Ops Last Status], "GREEN", "#2E7D32", "RED", "#C62828", "#B7791F" )
Ops Last Ready = VAR d = MAX ( Daily_Runs[Run_Date] ) VAR t = CALCULATE ( SELECTEDVALUE ( Daily_Runs[Ready_Time] ), Daily_Runs[Run_Date] = d ) RETURN IF ( ISBLANK ( t ), "not ready", t )
Ops Last Run Day = FORMAT ( MAX ( Daily_Runs[Run_Date] ), "ddd dd mmm" )
Ops SLA Met % = DIVIDE ( CALCULATE ( COUNTROWS ( Daily_Runs ), Daily_Runs[SLA_Met] = "Yes" ), CALCULATE ( COUNTROWS ( Daily_Runs ), Daily_Runs[SLA_Met] IN { "Yes", "No" } ) )
Ops Days SLA Missed = CALCULATE ( COUNTROWS ( Daily_Runs ), Daily_Runs[SLA_Met] = "No" ) + 0
Ops SLA Target = MAX ( Daily_Runs[SLA_Target] )
Run_Month = DATE ( YEAR ( Daily_Runs[Run_Date] ), MONTH ( Daily_Runs[Run_Date] ), 1 )   // calculated column on Daily_Runs, format "mmm yyyy": the axis of the SLA history chart
Ops SLA Color = SWITCH ( SELECTEDVALUE ( Daily_Runs[SLA_Met] ), "Yes", "#2E7D32", "No", "#C62828", "#605E5C" )
Ops Files OK Text = VAR d = MAX ( Daily_Runs[Run_Date] ) RETURN CALCULATE ( SUM ( Daily_Runs[Files_OK] ), Daily_Runs[Run_Date] = d ) & " of " & CALCULATE ( SUM ( Daily_Runs[Files_Expected] ), Daily_Runs[Run_Date] = d )
Ops Files Missing Today = VAR d = MAX ( Daily_Runs[Run_Date] ) RETURN CALCULATE ( SUM ( Daily_Runs[Files_Missing] ), Daily_Runs[Run_Date] = d ) + 0
Ops Files Failed Today = VAR d = MAX ( Daily_Runs[Run_Date] ) RETURN CALCULATE ( SUM ( Daily_Runs[Files_Failed] ), Daily_Runs[Run_Date] = d ) + 0
Ops Issues To Fix = CALCULATE ( COUNTROWS ( Data_Issues ), Data_Issues[Severity] = "WARN" ) + 0
Ops Issues Info = CALCULATE ( COUNTROWS ( Data_Issues ), Data_Issues[Severity] = "INFO" ) + 0
Ops Issue Count = COUNTROWS ( Data_Issues )
Ops Fix Where = SELECTEDVALUE ( Data_Issues[Fix_Where] )
Ops Oldest Load Days = MAX ( Data_Freshness[Days_Since_Load] )
Ops Oldest Source = VAR m = MAX ( Data_Freshness[Days_Since_Load] ) RETURN IF ( m = 0, "all loaded today", CONCATENATEX ( FILTER ( Data_Freshness, Data_Freshness[Days_Since_Load] = m ), Data_Freshness[Source], ", " ) )
```

## Catalogue: first build phase

### _Bookings

| Measure | Definition | Format |
| --- | --- | --- |
| Total Bookings | Distinct count of Opty Id | Whole number |
| Sales Closed | Bookings at stage 06 (Is_Sale_Closed = 1); shows 0 instead of blank | Whole number |
| Open Bookings | Bookings at stage 03 Open Yellow / Booked | Whole number |
| Vehicle Allotted | Bookings at stage 04 | Whole number |
| Allotted or Beyond | Bookings at stage 04 or higher; shows 0 instead of blank | Whole number |
| Active Advisors | Distinct count of Employee Full Name | Whole number |
| Bookings per Advisor | Total Bookings / Active Advisors | Decimal 1 |
| Closure % | Sales Closed / Total Bookings | Percent 1 |
| Allotment % | Allotted or Beyond / Total Bookings | Percent 1 |
| Test Drive % | Bookings with Testdrive = Given / Total Bookings | Percent 1 |
| Same Day Booking % | Bookings with Lead_to_Booking_Days = 0 / Total Bookings | Percent 1 |
| Avg Days Enquiry to Booking | Average of Lead_to_Booking_Days (YF Date minus Opty Created Date) | Decimal 1 |
| Stale Open Bookings | Stage 03 bookings older than 15 days vs latest data date; shows 0 instead of blank | Whole number |
| Source Share % | Bookings of a source / bookings of all sources | Percent 1 |
| Latest Data Date | Latest YF Date in the data (ignores filters) | Date |
| Bookings Latest Day | Total Bookings on the Latest Data Date | Whole number |
| Bookings MTD | Total Bookings from the 1st of the latest month to the latest date | Whole number |
| Bookings Prev Month Same Period | Bookings in the same day range of the previous month | Whole number |
| MoM % | (MTD - previous month same period) / previous month same period | Percent 1 |
| MoM Display | MoM % as text, or 'n/a - needs last month' when blank | Text |
| Projected Month End | MTD / day of month x days in month (simple run-rate) | Decimal 1 |
| Period End Date | Last day of the selected month, capped at the latest data date (All = latest data date) | Date |
| Bookings Month To Date | Bookings from the 1st of the selected month to Period End Date | Whole number |
| Bookings Same Days Last Month | Bookings in the same day range of the previous month | Whole number |
| Bookings Change vs Last Month % | (Month to date - same days last month) / same days last month | Percent 1 |
| Bookings Change Text | Arrow + % text for the card, or 'No data for last month yet' | Text |
| Bookings Change Colour | Green when up, red when down, grey when no comparison | Text |
| Daily Average Bookings | Month to date / days elapsed | Decimal 1 |
| Days Left In Month | Days remaining after Period End Date | Whole number |
| Projected Bookings Month End | Month to date / days elapsed x days in month (run rate) | Whole number |
| Sales Closed Month To Date | Stage 06 bookings made this month | Whole number |
| Closure Rate Month To Date | Sales Closed Month To Date / Bookings Month To Date | Percent 1 |
| Lost Bookings | Bookings at stage 07 Closed Lost | Whole number |
| Loss Rate % | Lost Bookings / Total Bookings | Percent 1 |
| Lost Bookings Month To Date | Stage 07 bookings made this month | Whole number |
| Cumulative Bookings This Month | Running total by day of month for the selected month | Whole number |
| Cumulative Bookings Last Month | Running total by day of month for the previous month | Whole number |

### _DataQuality

| Measure | Definition | Format |
| --- | --- | --- |
| DQ Issues | Number of issue rows (shows 0 instead of blank) | Whole number |
| DQ Warnings | Issue rows with Severity = WARN (shows 0 instead of blank) | Whole number |
| Records With Warnings | Distinct Opty Id with at least one WARN; shows 0 instead of blank | Whole number |
| Clean Records % | (Total Bookings - Records With Warnings) / Total Bookings | Percent 1 |
| Repeated Name Entries | Bookings where Name_Repeated_Flag = Yes (customer name typed twice) | Whole number |
| Missing From Latest | Bookings flagged Missing_From_Latest = Yes (booked this month, absent from latest file) | Whole number |
| Backward Moves | Stage changes where the stage number went down | Whole number |
| Possible Duplicate Bookings | Bookings with Repeat_Type = 'Possible duplicate - same variant' (same mobile + same variant, different Opty Id); counts every booking in the pair; shows 0 instead of blank | Whole number |
| Bookings That May Be Counted Twice | Sum of Dup_Extra: possible-duplicate bookings minus the oldest one of each group = how much Total Bookings may be overstated | Whole number |
| Model or Variant Changes | Bookings sharing a mobile with another Opty Id but with a different variant or model (customer changed his mind); the older Opty Id may still be open | Whole number |
| DQ Info | Issue rows with Severity = INFO | Whole number |
| Unspecified Advisor Bookings | Issue rows where Check = Advisor is 'Unspecified' | Whole number |
| Severity Colour | Hex colour for the Severity cell in the fix list (WARN amber, INFO grey) | Text |
| Hygiene Traffic Light | Hex colour from Clean Records %: >=95% green, >=85% amber, below red (edit the numbers inside) | Text |

### _PipelineOps (first SLA page; Data Operations now uses the Ops measures)

| Measure | Definition | Format |
| --- | --- | --- |
| Rows Loaded | Sum of Rows_Read | Whole number |
| New Bookings Loaded | Sum of New | Whole number |
| Updated Bookings | Sum of Updated | Whole number |
| Last Load Date | Latest Load_Date | Date |
| Days Since Last Load | Today minus Last Load Date | Whole number |
| Data Ready Time | Per day: the later of PV and EV first-load times (when BOTH were ready). Reruns and tests later in the day do not change it | Time |
| SLA Target | 13:00 | Time |
| SLA Days Total | Distinct Load_Date | Whole number |
| SLA Days Met | Days where Data Ready Time <= SLA Target | Whole number |
| SLA Adherence % | SLA Days Met / SLA Days Total | Percent 1 |
| Last Day Ready Time | Data Ready Time on the most recent load date | Time |
| Data Ready Hour | Data Ready Time as decimal hours (9.5 = 09:30) so it can be charted | Decimal 1 |
| SLA Target Hour | SLA Target as decimal hours (13 = 13:00), used as the target line | Decimal 1 |

### _Pipeline

| Measure | Definition | Format |
| --- | --- | --- |
| Snapshot Date | Date of today's open-pipeline snapshot | Date |
| Open Enquiries | Open Green / Pink Forms (current FY) | Whole number |
| Booked Awaiting Vehicle | Stage 03 Yellow Form still open | Whole number |
| Allotted Awaiting Delivery | Stage 04 Vehicle Allotted still open | Whole number |
| Booked Not Delivered | Booked Awaiting Vehicle + Allotted Awaiting Delivery (all months of the FY) | Whole number |
| Waiting Over 30 Days | Booked or allotted customers waiting more than 30 days since YF date | Whole number |
| Stale Enquiries | Enquiries open more than 30 days since Opty Created Date | Whole number |
| New Enquiries Last 7 Days | Enquiries created in the last 7 days | Whole number |
| Median Wait Days | Median days since YF date for booked/allotted customers | Decimal 0 |
| Hot Or Warm Enquiries | Open enquiries classified Hot or Warm | Whole number |
| Lead Class Blank % | Share of open records with lead classification Unspecified | Percent 1 |
| Left Pipeline Last Day | Records that left the pipeline on the latest snapshot date | Whole number |
| Pipeline Open Items | Every open item (enquiry, booked, allotted); used for charts and 'greater than 0' visual filters | Whole number |
| Pipeline Enquiries | Open Green/Pink Form enquiries (Form_Order 1-2); 0 instead of blank | Whole number |
| Pipeline New Enquiries 7 Days | Enquiries created in the last 7 days | Whole number |
| Pipeline Stale Enquiries | Enquiries open more than 30 days (Is_Stale_Enquiry = 1) | Whole number |
| Pipeline Hot Warm | Enquiries with lead class Hot or Warm | Whole number |
| Pipeline Hot | Enquiries with lead class Hot | Whole number |
| Pipeline Lead Not Set % | Share of enquiries with lead class Unspecified (all ages) | Percent 0 |
| Pipeline Test Drives Given | Enquiries with Testdrive = Given | Whole number |
| Pipeline Test Drives Not Given | Pipeline Enquiries - Pipeline Test Drives Given | Whole number |
| Pipeline Test Drive % | Test drives given / enquiries | Percent 0 |
| Pipeline Booked Not Delivered | Yellow Form + Allotted still open (Form_Order 3-4) | Whole number |
| Pipeline Waiting Over 30 | Booked or allotted, more than 30 days since YF date (Is_Long_Wait = 1) | Whole number |
| Pipeline Typical Wait Days | Median days since booking for booked + allotted | Whole number |
| Pipeline Allotted | Vehicle allotted, not delivered (Form_Order 4) | Whole number |
| Pipeline Allotted Typical Days | Median days since booking, allotted only | Whole number |
| Pipeline Allotted Longest Days | Max days since booking, allotted only | Whole number |
| Pipeline Hot Warm No Test Drive | Follow_Up_Order = 3: Hot/Warm enquiry up to 30 days old, no test drive | Whole number |
| Pipeline Lead Class Not Set | Follow_Up_Order = 4: enquiry up to 30 days old, lead class Unspecified | Whole number |
| Pipeline Snapshot Text | 'Pipeline as on dd mmm yyyy' from the latest Snapshot_Date (ignores filters) | Text |
| Follow Up Advice | Count + what-to-do sentence for the selected follow-up list | Text |

### _Team

| Measure | Definition | Format |
| --- | --- | --- |
| Team Open Enquiries | [Pipeline Enquiries] ignoring the Month filter | Whole number |
| Team Test Drive % | [Pipeline Test Drive %] ignoring the Month filter | Percent 0 |
| Team Hot Warm | [Pipeline Hot Warm] ignoring the Month filter | Whole number |
| Team Lead Not Set % | [Pipeline Lead Not Set %] ignoring the Month filter | Percent 0 |
| Team Stale Enquiries | [Pipeline Stale Enquiries] ignoring the Month filter | Whole number |
| Team Booked Not Delivered | [Pipeline Booked Not Delivered] ignoring the Month filter | Whole number |
| Team Waiting Over 30 | [Pipeline Waiting Over 30] ignoring the Month filter | Whole number |
| Team Activity | Bookings this month + open pipeline items; only used as a 'greater than 0' visual filter | Whole number |
| Active CAs | CAs with at least one booking in the selected month | Whole number |
| CAs In Team | CAs in Team_Mapping.xlsx for the selected SM / TL | Whole number |
| CAs With No Booking | CAs In Team minus Active CAs | Whole number |
| Bookings per Active CA | Bookings Month To Date / Active CAs | Decimal 1 |
| Cross Sell Bookings MTD | Bookings this month where the vehicle type differs from the CA's Team Type | Whole number |
| EV Team Sold PV MTD | Cross-sell bookings by EV-team CAs (they sold PV) | Whole number |
| PV Team Sold EV MTD | Cross-sell bookings by PV-team CAs (they sold EV) | Whole number |
