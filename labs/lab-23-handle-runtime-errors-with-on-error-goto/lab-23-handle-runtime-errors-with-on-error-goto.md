# Lab 23 — Handle Runtime Errors with On Error GoTo

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to identify relevant controls by trapping and responding to runtime errors.

## Goal
An analyst typing 'ten' instead of 10 must not be dumped into the code window. You will add an error handler that catches the failure, identifies it via Err.Number and shows a helpful message.

## What you'll build
A CheckError Sub with a labelled error handler that traps division-by-zero and bad input.

**Tools used:** On Error GoTo, error label, Err.Number, Select Case, Exit Sub

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-23-start.xlsm`](lab-23-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-23-end.xlsm`](lab-23-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 22. If you already have your own workbook from Lab 22, keep using it — `lab-23-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Prompts the user for two numbers; test with 10 and 0, and with 10 and 'ten'.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 23: Handle Runtime Errors with On Error GoTo.

Business context: An analyst typing 'ten' instead of 10 must not be dumped into the code window. You will add an error handler that catches the failure, identifies it via Err.Number and shows a helpful message.
Workbook goal: A CheckError Sub with a labelled error handler that traps division-by-zero and bad input.
Available workbook context: Prompts the user for two numbers; test with 10 and 0, and with 10 and 'ten'.
Required VBA concepts/tools: On Error GoTo, error label, Err.Number, Select Case, Exit Sub
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Dividing by zero shows 'You attempted to divide by zero!' instead of Excel's runtime error dialog, and the macro ends cleanly.

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
- [ ] Evidence: the code passes this acceptance test — Dividing by zero shows 'You attempted to divide by zero!' instead of Excel's runtime error dialog, and the macro ends cleanly.

### Refine with evidence

```text
The VBA did not yet pass Lab 23. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Dividing by zero shows 'You attempted to divide by zero!' instead of Excel's runtime error dialog, and the macro ends cleanly.
```

## Step-by-step

1. Write a Sub named CheckError declaring a, b As Integer and c As Double.
2. Add On Error GoTo ErrorHandler: as the first executable line of the procedure.
3. Prompt for two numbers with InputBox and assign them to a and b.
4. Divide a by b, store the result in c, and display it with MsgBox.
5. Add Exit Sub before the handler so successful runs do not fall into the error code.
6. Add the ErrorHandler: label followed by Select Case Err.Number.
7. Add Case 11 for divide-by-zero with the message 'You attempted to divide by zero!', plus a Case Else for anything unexpected.
8. Test with 10 and 2, then 10 and 0, then 10 and the text 'ten', confirming a friendly message every time.

## The code

```vba
Sub CheckError()

    Dim a, b As Integer
    Dim c As Double

    On Error GoTo ErrorHandler:

    a = InputBox("Enter the first number.")
    b = InputBox("Enter the second number.")
    c = a / b
    MsgBox c

    Exit Sub                        ' success must not fall into the handler

ErrorHandler:
    Select Case Err.Number
        Case 11                     ' Divide by zero error
            MsgBox "You attempted to divide by zero!"
        Case 13                     ' Type mismatch - user typed text
            MsgBox "Please enter a number, not text."
        Case Else
            MsgBox "Unexpected error " & Err.Number & ": " & Err.Description
    End Select

End Sub
```

## Test it
Dividing by zero shows 'You attempted to divide by zero!' instead of Excel's runtime error dialog, and the macro ends cleanly.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-23-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*