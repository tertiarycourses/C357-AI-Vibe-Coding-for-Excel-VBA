# Lab 9 — Branch with If…Then…ElseIf…Else

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to control execution flow with conditional statements.

## Goal
NorthWind pays banded commission: 8% above $10,000, 6% above $1,000, and 5% below that. You will implement the bands with If…Then…ElseIf…Else — the exact logic the capstone builds on.

## What you'll build
An ElseIfCommission Sub that applies the correct banded rate to the selected sale.

**Tools used:** If…Then, ElseIf, Else, End If, Round

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-09-start.xlsm`](lab-09-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-09-end.xlsm`](lab-09-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 8. If you already have your own workbook from Lab 8, keep using it — `lab-09-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet — the mock data deliberately spans all three commission bands.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 9: Branch with If…Then…ElseIf…Else.

Business context: NorthWind pays banded commission: 8% above $10,000, 6% above $1,000, and 5% below that. You will implement the bands with If…Then…ElseIf…Else — the exact logic the capstone builds on.
Workbook goal: An ElseIfCommission Sub that applies the correct banded rate to the selected sale.
Available workbook context: Uses the Sales sheet — the mock data deliberately spans all three commission bands.
Required VBA concepts/tools: If…Then, ElseIf, Else, End If, Round
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: $22,300 yields $1,784.00 at 8%, $5,300 yields $318.00 at 6%, and $980 yields $49.00 at 5%.

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
- [ ] Evidence: the code passes this acceptance test — $22,300 yields $1,784.00 at 8%, $5,300 yields $318.00 at 6%, and $980 yields $49.00 at 5%.

### Refine with evidence

```text
The VBA did not yet pass Lab 9. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: $22,300 yields $1,784.00 at 8%, $5,300 yields $318.00 at 6%, and $980 yields $49.00 at 5%.
```

## Step-by-step

1. Select a cell in the Sales Amount column on the Sales sheet.
2. Write the ElseIfCommission Sub declaring curCommission As Currency.
3. Add the first test: If ActiveCell.Value > 10000 Then apply an 8% rate.
4. Add an ElseIf branch for values above 1000 applying 6%.
5. Add a final Else branch applying 5%, then close the block with End If.
6. Wrap each calculation in Round(..., 2) so the currency value is sensible.
7. Display the band and commission in a MsgBox, then run it on Farah Ismail's $22,300 sale.
8. Re-run on Elson Goh's $980 and Jasmine Lee's $5,300 sales to prove all three branches work.

## The code

```vba
Sub ElseIfCommission()

    Dim curSale As Currency
    Dim curCommission As Currency
    Dim strBand As String

    curSale = ActiveCell.Value

    If curSale > 10000 Then
        curCommission = Round(curSale * 0.08, 2)
        strBand = "8%"
    ElseIf curSale > 1000 Then
        curCommission = Round(curSale * 0.06, 2)
        strBand = "6%"
    Else
        curCommission = Round(curSale * 0.05, 2)
        strBand = "5%"
    End If

    MsgBox "Band: " & strBand & Chr(10) & _
           "Commission: " & Format(curCommission, "$#,##0.00")

End Sub
```

## Test it
$22,300 yields $1,784.00 at 8%, $5,300 yields $318.00 at 6%, and $980 yields $49.00 at 5%.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-09-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*