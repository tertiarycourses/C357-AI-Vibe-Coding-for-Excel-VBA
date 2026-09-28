# Lab 8 — Build Expressions with VBA Operators

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to apply arithmetic, comparison, logical and concatenation operators.

## Goal
A commission engine is nothing but expressions. You will use all four operator families to compute a rep's gross commission, compare it against a bonus threshold, combine two conditions, and build a readable summary line for the manager's report.

## What you'll build
A SummariseRep Sub that computes figures and concatenates a one-line summary sentence.

**Tools used:** Arithmetic (+ - * /), comparison (> < =), logical (And Or Not), concatenation (&)

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-08-start.xlsm`](lab-08-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-08-end.xlsm`](lab-08-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 7. If you already have your own workbook from Lab 7, keep using it — `lab-08-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Mock data: quarterly targets per rep so comparisons have something realistic to test against.

Dataset: [`../data/northwind-targets.csv`](../data/northwind-targets.csv)

| RepID | Rep Name | Q1 Sales | Q1 Target | Years of Service |
| --- | --- | --- | --- | --- |
| R001 | Aisha Rahman | 12400 | 10000 | 4 |
| R002 | Brian Tan | 8600 | 10000 | 2 |
| R004 | Devi Kumar | 15900 | 12000 | 6 |
| R006 | Farah Ismail | 22300 | 20000 | 3 |
| R008 | Hui Ling Ong | 640 | 5000 | 1 |

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 8: Build Expressions with VBA Operators.

Business context: A commission engine is nothing but expressions. You will use all four operator families to compute a rep's gross commission, compare it against a bonus threshold, combine two conditions, and build a readable summary line for the manager's report.
Workbook goal: A SummariseRep Sub that computes figures and concatenates a one-line summary sentence.
Available workbook context: Mock data: quarterly targets per rep so comparisons have something realistic to test against.
Required VBA concepts/tools: Arithmetic (+ - * /), comparison (> < =), logical (And Or Not), concatenation (&)
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Aisha Rahman is reported as above target with a loyalty bonus; Hui Ling Ong is reported as below target.

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
- [ ] Evidence: the code passes this acceptance test — Aisha Rahman is reported as above target with a loyalty bonus; Hui Ling Ong is reported as below target.

### Refine with evidence

```text
The VBA did not yet pass Lab 8. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Aisha Rahman is reported as above target with a loyalty bonus; Hui Ling Ong is reported as below target.
```

## Step-by-step

1. Add a sheet named 'Targets' and paste in the northwind-targets.csv mock data.
2. Write the SummariseRep Sub declaring variables for sales, target, commission and years of service.
3. Use arithmetic operators to compute commission as sales * 0.1 and variance as sales - target.
4. Use a comparison operator to test whether sales exceeded the target.
5. Use the And operator to test for a loyalty bonus: sales above target AND service of 3 years or more.
6. Use the & operator to concatenate the rep name, commission and verdict into one sentence.
7. Display the sentence with MsgBox and run it against Aisha Rahman's row.
8. Run it against Hui Ling Ong's row and confirm the below-target branch reports correctly.

## The code

```vba
Sub SummariseRep()

    Dim strRep As String
    Dim curSales As Currency, curTarget As Currency
    Dim curCommission As Currency, curVariance As Currency
    Dim intYears As Integer
    Dim strVerdict As String

    strRep = ActiveCell.Offset(0, -1).Value
    curSales = ActiveCell.Value
    curTarget = ActiveCell.Offset(0, 1).Value
    intYears = ActiveCell.Offset(0, 2).Value

    curCommission = curSales * 0.1          ' arithmetic
    curVariance = curSales - curTarget

    If curSales > curTarget And intYears >= 3 Then      ' logical AND
        strVerdict = "above target with a loyalty bonus"
    ElseIf curSales > curTarget Then                    ' comparison
        strVerdict = "above target"
    Else
        strVerdict = "below target"
    End If

    ' concatenation
    MsgBox strRep & " earned " & Format(curCommission, "$#,##0.00") & _
           " and is " & strVerdict & " (variance " & _
           Format(curVariance, "$#,##0.00") & ")."

End Sub
```

## Test it
Aisha Rahman is reported as above target with a loyalty bonus; Hui Ling Ong is reported as below target.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-08-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*