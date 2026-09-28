# Lab 5 — Record a Macro to Calculate Sales Tax

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to use the macro recorder to generate VBA and then read the code it produces.

## Goal
The fastest way to discover Excel's object model is to let Excel write the code for you. You will record the steps for calculating 9% GST on each NorthWind sale, then open the recorded macro and read what the recorder generated.

## What you'll build
A recorded Cal_SalesTax macro plus your annotated reading of the generated code.

**Tools used:** Macro Recorder, Developer ribbon, Module

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-05-start.xlsm`](lab-05-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-05-end.xlsm`](lab-05-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 4. If you already have your own workbook from Lab 4, keep using it — `lab-05-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Mock data: the January sales extended with a GST Rate column.

Dataset: [`../data/northwind-sales-tax.csv`](../data/northwind-sales-tax.csv)

| RepID | Rep Name | Sales Amount | GST Rate | GST Amount |
| --- | --- | --- | --- | --- |
| R001 | Aisha Rahman | 12400 | 0.09 |  |
| R002 | Brian Tan | 8600 | 0.09 |  |
| R003 | Chen Wei Ling | 4750 | 0.09 |  |
| R004 | Devi Kumar | 15900 | 0.09 |  |
| R005 | Elson Goh | 980 | 0.09 |  |
| R006 | Farah Ismail | 22300 | 0.09 |  |

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 5: Record a Macro to Calculate Sales Tax.

Business context: The fastest way to discover Excel's object model is to let Excel write the code for you. You will record the steps for calculating 9% GST on each NorthWind sale, then open the recorded macro and read what the recorder generated.
Workbook goal: A recorded Cal_SalesTax macro plus your annotated reading of the generated code.
Available workbook context: Mock data: the January sales extended with a GST Rate column.
Required VBA concepts/tools: Macro Recorder, Developer ribbon, Module
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Running Cal_SalesTax re-creates the GST formula in E2, giving 1116 for Aisha Rahman's $12,400 sale.

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
- [ ] Evidence: the code passes this acceptance test — Running Cal_SalesTax re-creates the GST formula in E2, giving 1116 for Aisha Rahman's $12,400 sale.

### Refine with evidence

```text
The VBA did not yet pass Lab 5. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Running Cal_SalesTax re-creates the GST formula in E2, giving 1116 for Aisha Rahman's $12,400 sale.
```

## Step-by-step

1. Add a new sheet named 'SalesTax' and paste in the northwind-sales-tax.csv mock data.
2. Click Developer → Record Macro; name it Cal_SalesTax and choose 'This Workbook' from the Store macro in list, then click OK.
3. With recording running, click cell E2 and enter the formula =C2*D2, then press Enter.
4. Click Developer → Stop Recording.
5. Press Alt+F11 and open the module the recorder created to read the generated code.
6. Identify the recorder's use of Range(...).Select and ActiveCell.FormulaR1C1 in the generated lines.
7. Add a comment above the procedure describing what it does and when to run it.
8. Return to Excel, delete the value in E2, and re-run Cal_SalesTax from the Macros dialog to confirm it rebuilds the formula.

## The code

```vba
' Generated by the macro recorder, then tidied and commented.
Sub Cal_SalesTax()
    ' Writes the GST formula (Sales Amount x GST Rate) into the active row
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
End Sub
```

## Test it
Running Cal_SalesTax re-creates the GST formula in E2, giving 1116 for Aisha Rahman's $12,400 sale.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-05-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*