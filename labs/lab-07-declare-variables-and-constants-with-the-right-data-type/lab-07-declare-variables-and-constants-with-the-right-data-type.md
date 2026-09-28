# Lab 7 — Declare Variables and Constants with the Right Data Type

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to declare variables and constants using appropriate data types.

## Goal
Hard-coded numbers scattered through a macro are impossible to maintain. You will rewrite the commission calculation using declared variables for the changing values and a Const for the GST rate that must never change, and see what happens when code tries to reassign a constant.

## What you'll build
A CalculateNetCommission Sub using Dim variables with correct types and a Const GST rate.

**Tools used:** Dim, Const, Option Explicit, Currency/Double/String/Integer types

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-07-start.xlsm`](lab-07-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-07-end.xlsm`](lab-07-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 6. If you already have your own workbook from Lab 6, keep using it — `lab-07-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet mock data; commission rate 10%, GST fixed at 9%.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 7: Declare Variables and Constants with the Right Data Type.

Business context: Hard-coded numbers scattered through a macro are impossible to maintain. You will rewrite the commission calculation using declared variables for the changing values and a Const for the GST rate that must never change, and see what happens when code tries to reassign a constant.
Workbook goal: A CalculateNetCommission Sub using Dim variables with correct types and a Const GST rate.
Available workbook context: Uses the Sales sheet mock data; commission rate 10%, GST fixed at 9%.
Required VBA concepts/tools: Dim, Const, Option Explicit, Currency/Double/String/Integer types
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: The macro reports commission and GST for the selected rep, and assigning to GST_RATE raises 'Assignment to constant not permitted'.

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
- [ ] Evidence: the code passes this acceptance test — The macro reports commission and GST for the selected rep, and assigning to GST_RATE raises 'Assignment to constant not permitted'.

### Refine with evidence

```text
The VBA did not yet pass Lab 7. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: The macro reports commission and GST for the selected rep, and assigning to GST_RATE raises 'Assignment to constant not permitted'.
```

## Step-by-step

1. In the VBA Editor, type Option Explicit at the very top of Module1, above every procedure.
2. Create the CalculateNetCommission Sub as shown in the Learner Guide.
3. Declare Dim curSales As Currency, Dim curCommission As Currency and Dim strRep As String.
4. Declare Const GST_RATE As Double = 0.09 to hold the tax rate that must not change.
5. Read the rep name and sales amount from the active row using ActiveCell and Offset.
6. Calculate the commission at 10%, then the GST payable on it, and display both in a MsgBox.
7. Add the line GST_RATE = 0.07 inside the procedure and run it — read the compile error Excel raises.
8. Delete that line, then run the macro again successfully and review the data-type tables in the Learner Guide.

## The code

```vba
Option Explicit

Sub CalculateNetCommission()

    Const GST_RATE As Double = 0.09      ' fixed by law - must not change

    Dim strRep As String
    Dim curSales As Currency
    Dim curCommission As Currency
    Dim curGST As Currency

    strRep = ActiveCell.Offset(0, -3).Value
    curSales = ActiveCell.Value

    curCommission = curSales * 0.1
    curGST = curCommission * GST_RATE

    MsgBox strRep & Chr(10) & _
           "Commission: " & Format(curCommission, "$#,##0.00") & Chr(10) & _
           "GST payable: " & Format(curGST, "$#,##0.00")

    ' GST_RATE = 0.07
    ' ^ uncomment to see "Assignment to constant not permitted"

End Sub
```

## Test it
The macro reports commission and GST for the selected rep, and assigning to GST_RATE raises 'Assignment to constant not permitted'.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-07-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*