# Lab 22 — Identify Syntax, Runtime and Logical Errors

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to identify the three types of programming error.

## Goal
Before you can handle errors you must recognise them. You will deliberately introduce one of each kind into the commission macro and observe how differently each one fails — the logical error is the dangerous one, because Excel reports no problem at all.

## What you'll build
Three annotated broken procedures demonstrating each error type and its symptom.

**Tools used:** VBA compiler, runtime error dialog, Debug

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-22-start.xlsm`](lab-22-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-22-end.xlsm`](lab-22-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 21. If you already have your own workbook from Lab 21, keep using it — `lab-22-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Commission sheet mock data.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 22: Identify Syntax, Runtime and Logical Errors.

Business context: Before you can handle errors you must recognise them. You will deliberately introduce one of each kind into the commission macro and observe how differently each one fails — the logical error is the dangerous one, because Excel reports no problem at all.
Workbook goal: Three annotated broken procedures demonstrating each error type and its symptom.
Available workbook context: Uses the Commission sheet mock data.
Required VBA concepts/tools: VBA compiler, runtime error dialog, Debug
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: You can state which error is caught at compile time, which stops execution mid-run, and which produces a silently wrong answer.

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
- [ ] Evidence: the code passes this acceptance test — You can state which error is caught at compile time, which stops execution mid-run, and which produces a silently wrong answer.

### Refine with evidence

```text
The VBA did not yet pass Lab 22. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: You can state which error is caught at compile time, which stops execution mid-run, and which produces a silently wrong answer.
```

## Step-by-step

1. Copy your working commission macro into a new procedure named ErrorDemo.
2. Introduce a syntax error by deleting a closing parenthesis, then press F5 and read the compile message.
3. Restore the parenthesis, then introduce a runtime error by dividing the sale by a variable holding zero.
4. Run it and read the 'Run-time error 11: Division by zero' dialog; click Debug to see the highlighted line.
5. Press the Reset button to stop the macro, then remove the division.
6. Introduce a logical error by changing the commission rate from 0.1 to 0.01.
7. Run the macro and note that it completes with NO error at all — but every figure is wrong by a factor of ten.
8. Record in the Learner Guide table how each error type is detected and why logical errors need testing.

## The code

```vba
' --- 1. SYNTAX ERROR: caught when you compile; the code will not run at all
Sub SyntaxErrorDemo()
    Dim x As String
    x = UCase("hello"        ' <- missing closing parenthesis
    MsgBox x
End Sub


' --- 2. RUNTIME ERROR: valid code that fails while running (error 11)
Sub RuntimeErrorDemo()
    Dim a As Integer, b As Integer
    a = 100
    b = 0
    MsgBox a / b              ' <- division by zero stops execution here
End Sub


' --- 3. LOGICAL ERROR: runs perfectly and reports nothing,
'        but the answer is WRONG
Sub LogicalErrorDemo()
    Dim curCommission As Currency
    curCommission = ActiveCell.Value * 0.01    ' <- should be 0.1, not 0.01
    MsgBox curCommission          ' every figure is 10x too small
End Sub
```

## Test it
You can state which error is caught at compile time, which stops execution mid-run, and which produces a silently wrong answer.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-22-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*