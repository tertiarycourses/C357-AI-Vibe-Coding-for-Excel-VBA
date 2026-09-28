# Lab 3 — Message Boxes and Button Choices

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to use MsgBox to display information and capture the user's button choice.

## Goal
Before the tool recalculates commissions it must ask the user to confirm, because the operation overwrites last month's figures. You will use MsgBox with the vbYesNoCancel button set and read back which button was clicked.

## What you'll build
A ConfirmRecalculate Sub that asks the user to proceed and reports the button they chose.

**Tools used:** MsgBox, vbYesNoCancel, vbYes/vbNo/vbCancel constants

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-03-start.xlsm`](lab-03-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-03-end.xlsm`](lab-03-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 2. If you already have your own workbook from Lab 2, keep using it — `lab-03-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses NorthWind-Commission.xlsm from Lab 2.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 3: Message Boxes and Button Choices.

Business context: Before the tool recalculates commissions it must ask the user to confirm, because the operation overwrites last month's figures. You will use MsgBox with the vbYesNoCancel button set and read back which button was clicked.
Workbook goal: A ConfirmRecalculate Sub that asks the user to proceed and reports the button they chose.
Available workbook context: Uses NorthWind-Commission.xlsm from Lab 2.
Required VBA concepts/tools: MsgBox, vbYesNoCancel, vbYes/vbNo/vbCancel constants
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Each of the three buttons produces a different, correct follow-up message — Yes, No and Cancel are distinguished.

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
- [ ] Evidence: the code passes this acceptance test — Each of the three buttons produces a different, correct follow-up message — Yes, No and Cancel are distinguished.

### Refine with evidence

```text
The VBA did not yet pass Lab 3. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Each of the three buttons produces a different, correct follow-up message — Yes, No and Cancel are distinguished.
```

## Step-by-step

1. Open the VBA Editor and select Module1.
2. Write the ConfirmRecalculate Sub shown in the Learner Guide, using MsgBox with the vbYesNoCancel argument.
3. Note the MsgBox syntax: MsgBox(prompt, buttons, title) — the buttons argument controls which buttons appear.
4. Store the result in a variable so the code can react to the user's answer.
5. Add an If…Then…Else block that reports 'Recalculating…', 'Cancelled by user' or 'Operation aborted'.
6. Run the macro and click Yes; confirm the correct branch executes.
7. Run it twice more, clicking No and then Cancel, confirming each branch.
8. Review the common button constants (vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo) in the Learner Guide table.

## The code

```vba
Sub ConfirmRecalculate()

    Dim intAnswer As Integer

    intAnswer = MsgBox("Recalculate all commissions?", _
                       vbYesNoCancel, "Confirm")

    If intAnswer = vbYes Then
        MsgBox "Recalculating..."
    ElseIf intAnswer = vbNo Then
        MsgBox "Cancelled by user."
    Else
        MsgBox "Operation aborted."
    End If

End Sub
```

## Test it
Each of the three buttons produces a different, correct follow-up message — Yes, No and Cancel are distinguished.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-03-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*