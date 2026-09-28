# Lab 10 — Simplify Banded Logic with Select Case

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to use Select Case as a clearer alternative to long ElseIf chains.

## Goal
The finance team adds a fourth commission band, and the ElseIf chain is getting hard to read. You will rewrite the same logic with Select Case — the structure the capstone asks you to extend.

## What you'll build
A SelectRate Sub implementing four commission bands with Select Case.

**Tools used:** Select Case, Case Is, Case Else

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-10-start.xlsm`](lab-10-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-10-end.xlsm`](lab-10-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 9. If you already have your own workbook from Lab 9, keep using it — `lab-10-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Bands: 8% ≥ $10,000 · 6% ≥ $1,000 · 5% ≥ $500 · 4% below $500.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 10: Simplify Banded Logic with Select Case.

Business context: The finance team adds a fourth commission band, and the ElseIf chain is getting hard to read. You will rewrite the same logic with Select Case — the structure the capstone asks you to extend.
Workbook goal: A SelectRate Sub implementing four commission bands with Select Case.
Available workbook context: Bands: 8% ≥ $10,000 · 6% ≥ $1,000 · 5% ≥ $500 · 4% below $500.
Required VBA concepts/tools: Select Case, Case Is, Case Else
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Devi Kumar's $15,900 returns $1,272.00 at 8%, and Hui Ling Ong's $640 returns $32.00 at 5%.

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
- [ ] Evidence: the code passes this acceptance test — Devi Kumar's $15,900 returns $1,272.00 at 8%, and Hui Ling Ong's $640 returns $32.00 at 5%.

### Refine with evidence

```text
The VBA did not yet pass Lab 10. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Devi Kumar's $15,900 returns $1,272.00 at 8%, and Hui Ling Ong's $640 returns $32.00 at 5%.
```

## Step-by-step

1. Add a sheet named 'Commission' and copy the rep names and sales amounts across from the Sales sheet.
2. Write the SelectRate Sub, declaring curTotal As Currency and curCommission As Currency.
3. Read the selected sale into curTotal with ActiveCell.Value.
4. Open the block with Select Case curTotal.
5. Add Case Is >= 10000 applying 8%, then Case Is >= 1000 applying 6%, then Case Is >= 500 applying 5%.
6. Add Case Else applying 4%, then close with End Select.
7. Write the result back to the Commission column beside the sale rather than only showing a MsgBox.
8. Run the macro against several rows and confirm each lands in the right band.

## The code

```vba
Sub SelectRate()

    Dim curTotal As Currency
    Dim curCommission As Currency

    curTotal = ActiveCell.Value

    Select Case curTotal
        Case Is >= 10000
            curCommission = curTotal * 0.08
        Case Is >= 1000
            curCommission = curTotal * 0.06
        Case Is >= 500
            curCommission = curTotal * 0.05
        Case Else
            curCommission = curTotal * 0.04
    End Select

    ' write the result beside the sale
    ActiveCell.Offset(0, 1).Value = Round(curCommission, 2)

End Sub
```

## Test it
Devi Kumar's $15,900 returns $1,272.00 at 8%, and Hui Ling Ong's $640 returns $32.00 at 5%.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-10-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*