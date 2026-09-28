# Lab 14 — Write a Reusable Function that Returns a Value

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 03:** Function and Sub Procedure  
**Learning outcome:** LO3

## Objective
This lab develops your ability to create a user-defined function that accepts arguments and returns a result.

## Goal
The 10% commission calculation is now duplicated in three macros. You will extract it into a single user-defined Function that takes a sales figure and returns the commission — then call it from a worksheet cell exactly like SUM or AVERAGE.

## What you'll build
A Cal_Commission function callable from both VBA and any worksheet cell.

**Tools used:** Function…End Function, arguments, return value, UDF in a cell

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-14-start.xlsm`](lab-14-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-14-end.xlsm`](lab-14-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 13. If you already have your own workbook from Lab 13, keep using it — `lab-14-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Commission sheet mock data from Topic 2.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 14: Write a Reusable Function that Returns a Value.

Business context: The 10% commission calculation is now duplicated in three macros. You will extract it into a single user-defined Function that takes a sales figure and returns the commission — then call it from a worksheet cell exactly like SUM or AVERAGE.
Workbook goal: A Cal_Commission function callable from both VBA and any worksheet cell.
Available workbook context: Uses the Commission sheet mock data from Topic 2.
Required VBA concepts/tools: Function…End Function, arguments, return value, UDF in a cell
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: =Cal_Commission(C2) in a worksheet cell returns 1240 for Aisha Rahman's $12,400 sale.

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
- [ ] Evidence: the code passes this acceptance test — =Cal_Commission(C2) in a worksheet cell returns 1240 for Aisha Rahman's $12,400 sale.

### Refine with evidence

```text
The VBA did not yet pass Lab 14. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: =Cal_Commission(C2) in a worksheet cell returns 1240 for Aisha Rahman's $12,400 sale.
```

## Step-by-step

1. In Module1, create a new procedure beginning Function Cal_Commission(curSales As Currency) As Currency.
2. Inside the function, assign the result to the function's own name: Cal_Commission = curSales * 0.1.
3. Close the procedure with End Function.
4. Return to Excel, click an empty cell on the Commission sheet, and type =Cal_Commission(C2).
5. Press Enter and confirm the cell shows 10% of the sale — your function now behaves like a built-in.
6. Fill the formula down the column to apply it to all 12 reps.
7. Write a short Sub that also calls Cal_Commission and shows the result in a MsgBox, proving one function serves both callers.
8. Add a comment above the function documenting its argument, its return value and its rate assumption.

## The code

```vba
Function Cal_Commission(curSales As Currency) As Currency
    ' Returns 10% commission on the sales figure supplied.
    ' Callable from VBA AND from any worksheet cell:  =Cal_Commission(C2)
    Cal_Commission = curSales * 0.1
End Function


Sub UseTheFunction()
    Dim curPay As Currency
    curPay = Cal_Commission(ActiveCell.Value)
    MsgBox "Commission: " & Format(curPay, "$#,##0.00")
End Sub
```

## Test it
=Cal_Commission(C2) in a worksheet cell returns 1240 for Aisha Rahman's $12,400 sale.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-14-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*