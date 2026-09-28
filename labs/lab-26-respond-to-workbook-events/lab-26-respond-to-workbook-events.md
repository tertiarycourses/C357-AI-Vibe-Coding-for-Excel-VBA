# Lab 26 — Respond to Workbook Events

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 05:** Create User Defined Forms  
**Learning outcome:** LO5

## Objective
This lab develops your ability to write event procedures that run automatically at workbook level.

## Goal
Analysts keep closing the commission workbook before entering the month's rate, leaving the report blank. You will use Workbook_Open and Workbook_BeforeClose to greet the user on open and block closing until the rate is filled in.

## What you'll build
ThisWorkbook event procedures that run on open and guard the close.

**Tools used:** ThisWorkbook, Workbook_Open, Workbook_BeforeClose, Cancel argument

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-26-start.xlsm`](lab-26-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-26-end.xlsm`](lab-26-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 25. If you already have your own workbook from Lab 25, keep using it — `lab-26-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Commission sheet; cell F5 holds the month's commission rate.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 26: Respond to Workbook Events.

Business context: Analysts keep closing the commission workbook before entering the month's rate, leaving the report blank. You will use Workbook_Open and Workbook_BeforeClose to greet the user on open and block closing until the rate is filled in.
Workbook goal: ThisWorkbook event procedures that run on open and guard the close.
Available workbook context: Uses the Commission sheet; cell F5 holds the month's commission rate.
Required VBA concepts/tools: ThisWorkbook, Workbook_Open, Workbook_BeforeClose, Cancel argument
Component placement: Place workbook-event procedures in ThisWorkbook and standard helper Subs in Module1.
Acceptance test: Opening the file greets the user, and closing with F5 empty is blocked with a prompt until a rate is entered.

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
- [ ] Evidence: the code passes this acceptance test — Opening the file greets the user, and closing with F5 empty is blocked with a prompt until a rate is entered.

### Refine with evidence

```text
The VBA did not yet pass Lab 26. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Opening the file greets the user, and closing with F5 empty is blocked with a prompt until a rate is entered.
```

## Step-by-step

1. In the VBA Editor's Project Explorer, double-click ThisWorkbook to open its code window.
2. From the left-hand dropdown above the code window choose Workbook, and from the right-hand dropdown choose Open.
3. Inside Workbook_Open, display a MsgBox welcoming the user to the commission tracker.
4. From the right-hand dropdown now choose BeforeClose to generate its stub.
5. Inside it, test whether Worksheets("Commission").Range("F5").Value = "".
6. If it is empty, show a message asking for the commission rate and set Cancel = True to stop the close.
7. Save, close and reopen the workbook to confirm the welcome message fires.
8. Clear cell F5, try to close the workbook, and confirm the close is blocked; enter a rate and confirm it now closes.

## The code

```vba
' --- Both procedures go in the ThisWorkbook code window ---

Private Sub Workbook_Open()
    MsgBox "Welcome to the NorthWind Commission Tracker"
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Block the close until the month's commission rate has been entered
    If Worksheets("Commission").Range("F5").Value = "" Then
        MsgBox "Please enter a commission rate in cell F5 " & _
               "before closing the workbook."
        Cancel = True
    End If

End Sub
```

## Test it
Opening the file greets the user, and closing with F5 empty is blocked with a prompt until a rate is entered.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-26-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*