# Lab 21 — Automate Workbooks and Worksheets

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to manage workbook and worksheet objects from code.

## Goal
Month-end means creating a fresh sheet, naming it and checking whether last month's file is already open. You will drive the Workbook and Worksheet objects to do all three.

## What you'll build
A MonthEndSetup Sub that adds and names a sheet, and reports on open workbooks.

**Tools used:** Workbooks.Add/Open/Close, Worksheets.Add, .Name, .Activate, .Visible

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-21-start.xlsm`](lab-21-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-21-end.xlsm`](lab-21-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 20. If you already have your own workbook from Lab 20, keep using it — `lab-21-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses NorthWind-Commission.xlsm as built so far.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 21: Automate Workbooks and Worksheets.

Business context: Month-end means creating a fresh sheet, naming it and checking whether last month's file is already open. You will drive the Workbook and Worksheet objects to do all three.
Workbook goal: A MonthEndSetup Sub that adds and names a sheet, and reports on open workbooks.
Available workbook context: Uses NorthWind-Commission.xlsm as built so far.
Required VBA concepts/tools: Workbooks.Add/Open/Close, Worksheets.Add, .Name, .Activate, .Visible
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: A February sheet is created and named, the sheet count is reported correctly, and hiding then showing the sheet works.

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
- [ ] Evidence: the code passes this acceptance test — A February sheet is created and named, the sheet count is reported correctly, and hiding then showing the sheet works.

### Refine with evidence

```text
The VBA did not yet pass Lab 21. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: A February sheet is created and named, the sheet count is reported correctly, and hiding then showing the sheet works.
```

## Step-by-step

1. Write a Sub named MonthEndSetup in Module1.
2. Add a new worksheet with Worksheets.Add and immediately name it with ActiveSheet.Name = "February".
3. Activate the Sales sheet by name using Worksheets("Sales").Activate.
4. Report the number of sheets in the workbook with Worksheets.Count in a MsgBox.
5. Report the active workbook's name using ActiveWorkbook.Name.
6. Hide the February sheet with Worksheets("February").Visible = False, then run the macro.
7. Set Visible back to True and confirm the sheet reappears.
8. Add a commented line showing how Workbooks.Open would open another file by path.

## The code

```vba
Sub MonthEndSetup()

    Dim wsNew As Worksheet

    Set wsNew = Worksheets.Add
    wsNew.Name = "February"

    Worksheets("Sales").Activate

    MsgBox "Workbook: " & ActiveWorkbook.Name & Chr(10) & _
           "Sheets in this workbook: " & Worksheets.Count

    Worksheets("February").Visible = False      ' hide it
    Worksheets("February").Visible = True       ' and show it again

    ' Workbooks.Open "C:\Data\commission.xlsm"   ' opening another file

End Sub
```

## Test it
A February sheet is created and named, the sheet count is reported correctly, and hiding then showing the sheet works.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-21-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*