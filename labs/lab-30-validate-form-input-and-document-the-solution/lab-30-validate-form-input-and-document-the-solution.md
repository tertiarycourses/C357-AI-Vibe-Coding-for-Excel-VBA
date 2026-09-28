# Lab 30 — Validate Form Input and Document the Solution

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to identify relevant controls for data quality and produce solution documentation.

## Goal
A blank name or a text sales figure must never reach the sheet. You will validate every field before writing, and then document the finished tool so a colleague can run and maintain it.

## What you'll build
A validated frmSalesEntry plus an in-workbook ReadMe documenting the solution.

**Tools used:** IsNumeric, Trim, SetFocus, comments, documentation sheet

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-30-start.xlsm`](lab-30-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-30-end.xlsm`](lab-30-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 29. If you already have your own workbook from Lab 29, keep using it — `lab-30-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Test cases: blank name, non-numeric sales amount, and a valid record.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 30: Validate Form Input and Document the Solution.

Business context: A blank name or a text sales figure must never reach the sheet. You will validate every field before writing, and then document the finished tool so a colleague can run and maintain it.
Workbook goal: A validated frmSalesEntry plus an in-workbook ReadMe documenting the solution.
Available workbook context: Test cases: blank name, non-numeric sales amount, and a valid record.
Required VBA concepts/tools: IsNumeric, Trim, SetFocus, comments, documentation sheet
Component placement: Place form event procedures in frmSalesEntry and standard launch/helper Subs in Module1.
Acceptance test: Blank names, non-numeric amounts and missing regions are each rejected with a clear message and no row is written.

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
- [ ] Evidence: the code passes this acceptance test — Blank names, non-numeric amounts and missing regions are each rejected with a clear message and no row is written.

### Refine with evidence

```text
The VBA did not yet pass Lab 30. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Blank names, non-numeric amounts and missing regions are each rejected with a clear message and no row is written.
```

## Step-by-step

1. In cmdSubmit_Click, test whether Trim(txtRepName.Value) = "" and warn the user if so.
2. Use txtRepName.SetFocus to return the cursor to the offending field, then Exit Sub to stop the write.
3. Add a second check using IsNumeric(txtSalesAmount.Value) to reject non-numeric sales figures.
4. Add a check that a region has been selected in cboRegion.
5. Test all three failure cases and confirm nothing is written to the sheet.
6. Add comments above each procedure stating its purpose, its inputs and its effect on the workbook.
7. Add a ReadMe worksheet listing what the tool does, how to launch the form and the commission bands used.
8. Save the workbook as .xlsm and confirm all macros survive the save.

## The code

```vba
Private Sub cmdSubmit_Click()

    Dim lngRow As Long

    ' ---- validate BEFORE writing anything to the sheet ----
    If Trim(txtRepName.Value) = "" Then
        MsgBox "Please enter the rep name."
        txtRepName.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtSalesAmount.Value) Then
        MsgBox "The sales amount must be a number."
        txtSalesAmount.SetFocus
        Exit Sub
    End If

    If cboRegion.Value = "" Then
        MsgBox "Please select a region."
        cboRegion.SetFocus
        Exit Sub
    End If

    ' ---- only now is it safe to write ----
    Worksheets("Sales").Activate
    lngRow = Worksheets("Sales").Range("A1048576").End(xlUp).Row + 1

    Cells(lngRow, 1) = txtRepID.Value
    Cells(lngRow, 2) = txtRepName.Value
    Cells(lngRow, 3) = cboRegion.Value
    Cells(lngRow, 5) = CCur(txtSalesAmount.Value)

    MsgBox "Record added to row " & lngRow & "."

End Sub
```

## Test it
Blank names, non-numeric amounts and missing regions are each rejected with a clear message and no row is written.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-30-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*