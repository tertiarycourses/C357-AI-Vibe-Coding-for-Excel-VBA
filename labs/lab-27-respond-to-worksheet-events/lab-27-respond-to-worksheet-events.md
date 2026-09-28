# Lab 27 — Respond to Worksheet Events

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to write event procedures that react to changes on a worksheet.

## Goal
Commission figures are generated, not typed — an analyst overtyping one silently corrupts the report. You will use Worksheet_Change to ask for confirmation whenever a cell on the Commission sheet is edited.

## What you'll build
A Worksheet_Change event on the Commission sheet that confirms every manual edit.

**Tools used:** Worksheet_Change, Target argument, EnableEvents, MsgBox vbOKCancel

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-27-start.xlsm`](lab-27-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-27-end.xlsm`](lab-27-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 26. If you already have your own workbook from Lab 26, keep using it — `lab-27-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Commission sheet mock data.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 27: Respond to Worksheet Events.

Business context: Commission figures are generated, not typed — an analyst overtyping one silently corrupts the report. You will use Worksheet_Change to ask for confirmation whenever a cell on the Commission sheet is edited.
Workbook goal: A Worksheet_Change event on the Commission sheet that confirms every manual edit.
Available workbook context: Uses the Commission sheet mock data.
Required VBA concepts/tools: Worksheet_Change, Target argument, EnableEvents, MsgBox vbOKCancel
Component placement: Place Worksheet event procedures in the Commission sheet code window.
Acceptance test: Editing any cell on the Commission sheet prompts for confirmation, and choosing Cancel restores the previous value.

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
- [ ] Evidence: the code passes this acceptance test — Editing any cell on the Commission sheet prompts for confirmation, and choosing Cancel restores the previous value.

### Refine with evidence

```text
The VBA did not yet pass Lab 27. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Editing any cell on the Commission sheet prompts for confirmation, and choosing Cancel restores the previous value.
```

## Step-by-step

1. Right-click the Commission sheet tab and choose View Code.
2. From the left dropdown choose Worksheet, and from the right dropdown choose Change.
3. Inside Worksheet_Change, declare intResponse As Integer.
4. Show a MsgBox with vbOKCancel warning that this is an automatically generated worksheet.
5. If the user clicks OK, confirm the change; if Cancel, undo it and report that the change was rejected.
6. Wrap the undo in Application.EnableEvents = False and True so the event does not re-trigger itself.
7. Return to Excel and type a value into a commission cell to trigger the event.
8. Test both the OK and the Cancel path, confirming Cancel restores the original value.

## The code

```vba
' --- This procedure goes in the Commission SHEET's code window ---
' (right-click the sheet tab -> View Code)

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim intResponse As Integer

    intResponse = MsgBox("This is an automatically generated " & _
                         "worksheet. " & _
                         "Are you sure you want to change it?", vbOKCancel)

    If intResponse = vbOK Then
        MsgBox "Worksheet value changed."
    Else
        ' stop the undo from re-triggering this same event
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        MsgBox "Change rejected."
    End If

End Sub
```

## Test it
Editing any cell on the Commission sheet prompts for confirmation, and choosing Cancel restores the previous value.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-27-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*