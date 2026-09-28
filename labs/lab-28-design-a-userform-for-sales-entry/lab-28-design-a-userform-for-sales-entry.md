# Lab 28 — Design a UserForm for Sales Entry

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to produce a user-defined form with appropriately named controls.

## Goal
Typing straight into cells causes bad data. You will design a proper data-entry dialog with labels, text boxes, a combo box for region and Submit/Cancel buttons, naming every control so the code that follows is readable.

## What you'll build
A frmSalesEntry UserForm laid out with named, captioned controls.

**Tools used:** Insert → UserForm, Toolbox, Label, TextBox, ComboBox, CommandButton, Properties window

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-28-start.xlsm`](lab-28-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-28-end.xlsm`](lab-28-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 27. If you already have your own workbook from Lab 27, keep using it — `lab-28-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Captures RepID, Rep Name, Region and Sales Amount to match the Sales sheet columns.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 28: Design a UserForm for Sales Entry.

Business context: Typing straight into cells causes bad data. You will design a proper data-entry dialog with labels, text boxes, a combo box for region and Submit/Cancel buttons, naming every control so the code that follows is readable.
Workbook goal: A frmSalesEntry UserForm laid out with named, captioned controls.
Available workbook context: Captures RepID, Rep Name, Region and Sales Amount to match the Sales sheet columns.
Required VBA concepts/tools: Insert → UserForm, Toolbox, Label, TextBox, ComboBox, CommandButton, Properties window
Component placement: Place form event procedures in frmSalesEntry and standard launch/helper Subs in Module1.
Acceptance test: Running frmSalesEntry shows a tidy dialog with four labelled inputs, a region dropdown and two buttons.

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
- [ ] Evidence: the code passes this acceptance test — Running frmSalesEntry shows a tidy dialog with four labelled inputs, a region dropdown and two buttons.

### Refine with evidence

```text
The VBA did not yet pass Lab 28. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Running frmSalesEntry shows a tidy dialog with four labelled inputs, a region dropdown and two buttons.
```

## Step-by-step

1. In the VBA Editor choose Insert → UserForm; a blank form and the Toolbox appear.
2. In the Properties window set the form's Name to frmSalesEntry and its Caption to 'NorthWind Sales Entry'.
3. Drag four Labels onto the form and caption them Rep ID, Rep Name, Region and Sales Amount.
4. Add three TextBoxes named txtRepID, txtRepName and txtSalesAmount beside their labels.
5. Add a ComboBox named cboRegion for the region choice.
6. Add two CommandButtons named cmdSubmit and cmdCancel, captioned Submit and Cancel.
7. Align the controls neatly and size the form so nothing is clipped.
8. Press F5 to run the form and confirm it displays, then close it with the window's X.

## The code

```vba
' Lab 28 is form DESIGN - done in the editor, not in code.
'
'   Insert -> UserForm
'   Properties window:  (Name) = frmSalesEntry
'                       Caption = NorthWind Sales Entry
'
'   Controls to add from the Toolbox, and the names to give them:
'     Label       "Rep ID"          Label       "Rep Name"
'     Label       "Region"          Label       "Sales Amount"
'     TextBox     txtRepID          TextBox     txtRepName
'     ComboBox    cboRegion         TextBox     txtSalesAmount
'     CommandButton cmdSubmit  (Caption: Submit)
'     CommandButton cmdCancel  (Caption: Cancel)
'
'   Press F5 to run the form and confirm it displays cleanly.
```

## Test it
Running frmSalesEntry shows a tidy dialog with four labelled inputs, a region dropdown and two buttons.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-28-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*