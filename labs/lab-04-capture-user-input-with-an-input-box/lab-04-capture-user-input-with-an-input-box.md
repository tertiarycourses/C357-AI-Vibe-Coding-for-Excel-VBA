# Lab 4 — Capture User Input with an Input Box

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to prompt the user for a value with InputBox and use it in a calculation.

## Goal
Commission rates change between quarters, so the tool must not hard-code them. You will use InputBox to ask the analyst for the rate, apply it to the selected sale, and display the result.

## What you'll build
A ShowCommission Sub that asks for a commission rate and reports the commission on the selected cell.

**Tools used:** InputBox, ActiveCell, MsgBox

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-04-start.xlsm`](lab-04-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-04-end.xlsm`](lab-04-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 3. If you already have your own workbook from Lab 3, keep using it — `lab-04-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet mock data from Lab 1.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 4: Capture User Input with an Input Box.

Business context: Commission rates change between quarters, so the tool must not hard-code them. You will use InputBox to ask the analyst for the rate, apply it to the selected sale, and display the result.
Workbook goal: A ShowCommission Sub that asks for a commission rate and reports the commission on the selected cell.
Available workbook context: Uses the Sales sheet mock data from Lab 1.
Required VBA concepts/tools: InputBox, ActiveCell, MsgBox
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Entering 0.1 with the $22,300 sale selected reports a commission of $2,230.00.

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
- [ ] Evidence: the code passes this acceptance test — Entering 0.1 with the $22,300 sale selected reports a commission of $2,230.00.

### Refine with evidence

```text
The VBA did not yet pass Lab 4. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Entering 0.1 with the $22,300 sale selected reports a commission of $2,230.00.
```

## Step-by-step

1. Open the Sales sheet and click a cell in the Sales Amount column, for example E2.
2. In Module1, write the ShowCommission Sub given in the Learner Guide.
3. Use InputBox to prompt 'What is the commission rate? (e.g. 0.1 for 10%)'.
4. Read the selected sale with ActiveCell.Value.
5. Multiply the sale by the rate and store it in a commission variable.
6. Display the commission in a MsgBox, formatted as currency for readability.
7. Run the macro with a rate of 0.1 against Farah Ismail's $22,300 sale.
8. Run it again and click Cancel on the InputBox to observe that it returns an empty string.

## The code

```vba
Sub ShowCommission()

    Dim curRate As Double
    Dim curSale As Currency
    Dim curCommission As Currency

    curRate = InputBox("What is the commission rate? (e.g. 0.1 for 10%)")
    curSale = ActiveCell.Value
    curCommission = curSale * curRate

    MsgBox "Commission: " & Format(curCommission, "$#,##0.00")

End Sub
```

## Test it
Entering 0.1 with the $22,300 sale selected reports a commission of $2,230.00.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-04-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*