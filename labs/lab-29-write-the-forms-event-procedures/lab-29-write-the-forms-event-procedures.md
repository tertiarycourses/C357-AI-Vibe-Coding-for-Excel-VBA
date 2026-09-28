# Lab 29 — Write the Form's Event Procedures

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to produce event procedures that initialise a form and write its data to the worksheet.

## Goal
A form that does nothing is decoration. You will fill the region list as the form opens, append the captured record to the first empty row on submit — reusing the Do loop from Lab 13 — and close the form on cancel.

## What you'll build
A working frmSalesEntry that appends each submitted record to the Sales sheet.

**Tools used:** UserForm_Initialize, CommandButton_Click, AddItem, End(xlUp).Row, Unload Me

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-29-start.xlsm`](lab-29-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-29-end.xlsm`](lab-29-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 28. If you already have your own workbook from Lab 28, keep using it — `lab-29-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Regions loaded into the combo box: Central, East, North, West.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 29: Write the Form's Event Procedures.

Business context: A form that does nothing is decoration. You will fill the region list as the form opens, append the captured record to the first empty row on submit — reusing the Do loop from Lab 13 — and close the form on cancel.
Workbook goal: A working frmSalesEntry that appends each submitted record to the Sales sheet.
Available workbook context: Regions loaded into the combo box: Central, East, North, West.
Required VBA concepts/tools: UserForm_Initialize, CommandButton_Click, AddItem, End(xlUp).Row, Unload Me
Component placement: Place form event procedures in frmSalesEntry and standard launch/helper Subs in Module1.
Acceptance test: Submitting three records appends three new rows to the Sales sheet below the existing data, and Cancel closes the form.

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
- [ ] Evidence: the code passes this acceptance test — Submitting three records appends three new rows to the Sales sheet below the existing data, and Cancel closes the form.

### Refine with evidence

```text
The VBA did not yet pass Lab 29. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Submitting three records appends three new rows to the Sales sheet below the existing data, and Cancel closes the form.
```

## Step-by-step

1. Right-click the form and choose View Code to open its code window.
2. Select UserForm from the left dropdown and Initialize from the right to create UserForm_Initialize.
3. Inside it, clear the text boxes and load the four regions into cboRegion with .AddItem.
4. Double-click the Submit button to create cmdSubmit_Click.
5. Find the next free row with Worksheets("Sales").Range("A1048576").End(xlUp).Row + 1.
6. Write each control's .Value into the matching column of that row using Cells(lngRow, n).
7. Clear the controls afterwards so the form is ready for the next entry, and confirm with a MsgBox.
8. Double-click the Cancel button and add Unload Me, then run the form and submit three test records.

## The code

```vba
' --- All of this goes in the frmSalesEntry code window ---

Private Sub UserForm_Initialize()

    ' Clear the inputs and load the region list as the form opens
    txtRepID.Value = ""
    txtRepName.Value = ""
    txtSalesAmount.Value = ""

    cboRegion.Clear
    cboRegion.AddItem "Central"
    cboRegion.AddItem "East"
    cboRegion.AddItem "North"
    cboRegion.AddItem "West"

End Sub


Private Sub cmdSubmit_Click()

    Dim lngRow As Long

    Worksheets("Sales").Activate

    ' first empty row below the existing data
    lngRow = Worksheets("Sales").Range("A1048576").End(xlUp).Row + 1

    Cells(lngRow, 1) = txtRepID.Value
    Cells(lngRow, 2) = txtRepName.Value
    Cells(lngRow, 3) = cboRegion.Value
    Cells(lngRow, 5) = txtSalesAmount.Value

    ' reset the form for the next entry
    txtRepID.Value = ""
    txtRepName.Value = ""
    txtSalesAmount.Value = ""
    cboRegion.Value = ""

    MsgBox "Record added to row " & lngRow & "."

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
```

## Test it
Submitting three records appends three new rows to the Sales sheet below the existing data, and Cancel closes the form.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-29-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*