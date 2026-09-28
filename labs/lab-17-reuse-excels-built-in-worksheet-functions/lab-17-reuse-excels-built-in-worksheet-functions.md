# Lab 17 — Reuse Excel's Built-in Worksheet Functions

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 03:** Function and Sub Procedure  
**Learning outcome:** LO3

## Objective
This lab develops your ability to call Excel's built-in functions from VBA with Application.WorksheetFunction.

## Goal
Do not re-implement what Excel already does. You will use Application.WorksheetFunction to total and summarise NorthWind's sales, and to compute a loan repayment with PMT for the company's new vehicle financing.

## What you'll build
A SalesStatistics Sub reporting total, average, highest and lowest sales, plus a PMT calculation.

**Tools used:** Application.WorksheetFunction (Sum, Average, Max, Min, Pmt), Range

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-17-start.xlsm`](lab-17-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-17-end.xlsm`](lab-17-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 16. If you already have your own workbook from Lab 16, keep using it — `lab-17-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Mock data: vehicle loan terms — $18,000 over 12 months at 6% annual interest.

Dataset: [`../data/northwind-loan.csv`](../data/northwind-loan.csv)

| Item | Value |
| --- | --- |
| Loan Principal | 18000 |
| Annual Interest Rate | 0.06 |
| Term (months) | 12 |

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 17: Reuse Excel's Built-in Worksheet Functions.

Business context: Do not re-implement what Excel already does. You will use Application.WorksheetFunction to total and summarise NorthWind's sales, and to compute a loan repayment with PMT for the company's new vehicle financing.
Workbook goal: A SalesStatistics Sub reporting total, average, highest and lowest sales, plus a PMT calculation.
Available workbook context: Mock data: vehicle loan terms — $18,000 over 12 months at 6% annual interest.
Required VBA concepts/tools: Application.WorksheetFunction (Sum, Average, Max, Min, Pmt), Range
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Total January sales report as $92,470, and the monthly loan repayment is approximately $1,549 — matching Excel's own =PMT result.

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
- [ ] Evidence: the code passes this acceptance test — Total January sales report as $92,470, and the monthly loan repayment is approximately $1,549 — matching Excel's own =PMT result.

### Refine with evidence

```text
The VBA did not yet pass Lab 17. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Total January sales report as $92,470, and the monthly loan repayment is approximately $1,549 — matching Excel's own =PMT result.
```

## Step-by-step

1. Add a sheet named 'Finance' and paste in the northwind-loan.csv mock data.
2. Write the SalesStatistics Sub and set a Range object over the Sales Amount column.
3. Use Application.WorksheetFunction.Sum on that range to total NorthWind's January sales.
4. Add Average, Max and Min calls over the same range.
5. Concatenate the four figures into a summary message and display it.
6. Write a second Sub, CalcPayment, declaring dRate, iPeriods and cPV variables.
7. Call Application.WorksheetFunction.Pmt(dRate / iPeriods, iPeriods, cPV) and multiply by -1 for a positive figure.
8. Display the monthly repayment and check it against Excel's own =PMT formula in a cell.

## The code

```vba
Sub SalesStatistics()

    Dim rngSales As Range
    Dim curTotal As Currency, curAverage As Currency
    Dim curMax As Currency, curMin As Currency

    Set rngSales = Worksheets("Sales").Range("E2:E13")

    curTotal = Application.WorksheetFunction.Sum(rngSales)
    curAverage = Application.WorksheetFunction.Average(rngSales)
    curMax = Application.WorksheetFunction.Max(rngSales)
    curMin = Application.WorksheetFunction.Min(rngSales)

    MsgBox "Total:   " & Format(curTotal, "$#,##0.00") & Chr(10) & _
           "Average: " & Format(curAverage, "$#,##0.00") & Chr(10) & _
           "Highest: " & Format(curMax, "$#,##0.00") & Chr(10) & _
           "Lowest:  " & Format(curMin, "$#,##0.00")

End Sub


Sub CalcPayment()

    Dim dRate As Double
    Dim iPeriods As Integer
    Dim cPV As Currency, cPayment As Currency

    dRate = 0.06            ' annual interest rate
    iPeriods = 12           ' months
    cPV = 18000             ' principal

    cPayment = Application.WorksheetFunction.Pmt( _
                   dRate / iPeriods, iPeriods, cPV) * -1

    MsgBox "Monthly repayment: " & Format(cPayment, "$#,##0.00")

End Sub
```

## Test it
Total January sales report as $92,470, and the monthly loan repayment is approximately $1,549 — matching Excel's own =PMT result.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-17-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*