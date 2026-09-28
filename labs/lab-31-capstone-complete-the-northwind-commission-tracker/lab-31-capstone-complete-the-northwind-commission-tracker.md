# Lab 31 — Capstone — Complete the NorthWind Commission Tracker

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to integrate the full course into one working automated solution.

## Goal
Bring every topic together into the finished deliverable: banded commission with Select Case, a launch button, the validated entry form, error handling and an on-demand chart.

## What you'll build
The completed NorthWind-Commission.xlsm automating capture, calculation, validation and reporting.

**Tools used:** Select Case, UserForm, error handling, charts, Form Controls, .xlsm

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-31-start.xlsm`](lab-31-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-31-end.xlsm`](lab-31-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 30. If you already have your own workbook from Lab 30, keep using it — `lab-31-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Final bands: 8% ≥ $10,000 · 7% ≥ $5,000 · 6% ≥ $1,000 · 5% ≥ $500 · 4% below $500.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 31: Capstone — Complete the NorthWind Commission Tracker.

Business context: Bring every topic together into the finished deliverable: banded commission with Select Case, a launch button, the validated entry form, error handling and an on-demand chart.
Workbook goal: The completed NorthWind-Commission.xlsm automating capture, calculation, validation and reporting.
Available workbook context: Final bands: 8% ≥ $10,000 · 7% ≥ $5,000 · 6% ≥ $1,000 · 5% ≥ $500 · 4% below $500.
Required VBA concepts/tools: Select Case, UserForm, error handling, charts, Form Controls, .xlsm
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: From a fresh open: the welcome fires, 'New Sale' captures a validated record, commission is banded correctly including the 7% band, 'Plot Commission' charts the result, and the file reopens with everything intact.

Requirements:
- Start standard modules with Option Explicit.
- Use the exact worksheet, range, control and procedure names stated above; do not invent objects.
- Declare every variable with an appropriate data type.
- Keep the procedure focused on this lab and add short comments explaining intent.
- Qualify Workbook, Worksheet and Range references where practical.
- Add safe error handling when the requested operation can fail.
- Do not delete files, call the shell, access credentials, use the network or hide code.
- Return: (1) where the code belongs, (2) one complete VBA code block, and
  (3) three manual test cases including the acceptance test.
```

### Review before you run

- [ ] Placement: the code is in the correct Module, ThisWorkbook, worksheet or UserForm component.
- [ ] Names: every worksheet, range, control and procedure name exists in the supplied workbook.
- [ ] Logic: the business rule and boundary cases match the lab goal; no unexplained magic values remain.
- [ ] Safety: there are no unexpected file, shell, network, credential, deletion or hidden-code operations.
- [ ] Quality: variables are declared, the code is readable, and errors are handled where appropriate.
- [ ] Evidence: the code passes this acceptance test — From a fresh open: the welcome fires, 'New Sale' captures a validated record, commission is banded correctly including the 7% band, 'Plot Commission' charts the result, and the file reopens with everything intact.

### Refine with evidence

```text
The VBA did not yet pass Lab 31. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: From a fresh open: the welcome fires, 'New Sale' captures a validated record, commission is banded correctly including the 7% band, 'Plot Commission' charts the result, and the file reopens with everything intact.
```

## Step-by-step

1. Open your NorthWind-Commission.xlsm and review the Sales, Commission, Summary and ReadMe sheets.
2. Modify the SelectRate procedure from Lab 10 to add the new 7% band for sales of $5,000 or more.
3. Re-run SelectRate across the data and confirm Jasmine Lee's $5,300 sale now earns 7%, not 6%.
4. Add a Form Control button on the Sales sheet labelled 'New Sale' that shows frmSalesEntry.
5. Use the form to enter three new sales records and confirm they append correctly.
6. Confirm the 'Plot Commission' button from Lab 18 charts the updated commission figures.
7. Wrap the master procedure in On Error GoTo so any unexpected failure reports cleanly.
8. Save as .xlsm, close and reopen the file, and confirm every button, event and macro still works.

## The code

```vba
' ===== CAPSTONE: the finished NorthWind Commission Tracker =====

' 1. Banded commission, now including the new 7% band
Sub SelectRate()

    Dim curTotal As Currency
    Dim curCommission As Currency

    On Error GoTo ErrorHandler:

    curTotal = ActiveCell.Value

    Select Case curTotal
        Case Is >= 10000
            curCommission = curTotal * 0.08
        Case Is >= 5000                      ' <- the new band
            curCommission = curTotal * 0.07
        Case Is >= 1000
            curCommission = curTotal * 0.06
        Case Is >= 500
            curCommission = curTotal * 0.05
        Case Else
            curCommission = curTotal * 0.04
    End Select

    ActiveCell.Offset(0, 1).Value = Round(curCommission, 2)
    Exit Sub

ErrorHandler:
    MsgBox "Could not calculate commission: " & Err.Description

End Sub


' 2. The 'New Sale' button opens the validated entry form
Sub ShowSalesForm()
    frmSalesEntry.Show
End Sub


' 3. The 'Plot Commission' button charts the updated figures
Sub PlotCommission()
    Worksheets("Commission").Activate
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range("B1:B13,D1:D13")
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "NorthWind Commission by Rep"
End Sub
```

## Test it
From a fresh open: the welcome fires, 'New Sale' captures a validated record, commission is banded correctly including the 7% band, 'Plot Commission' charts the result, and the file reopens with everything intact.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-31-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*