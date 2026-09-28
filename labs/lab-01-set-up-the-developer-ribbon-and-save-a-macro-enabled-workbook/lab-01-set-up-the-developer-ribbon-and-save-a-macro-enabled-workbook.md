# Lab 1 — Set Up the Developer Ribbon and Save a Macro-Enabled Workbook

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to access the VBA development environment and save work in the correct macro-enabled format.

## Goal
You have just joined NorthWind Trading as a sales analyst. Before you can automate anything, you must switch on Excel's Developer tools, open the Visual Basic Editor, and learn why a workbook containing macros has to be saved as .xlsm.

## What you'll build
A macro-enabled workbook, NorthWind-Commission.xlsm, with the Developer ribbon enabled and the VBA Editor open.

**Tools used:** Excel, Developer ribbon, Visual Basic Editor

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-01-start.xlsm`](lab-01-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-01-end.xlsm`](lab-01-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> Lab 1 starts from a blank workbook, which is exactly what `lab-01-start.xlsm` contains.

## Mock data
Mock data: 12 NorthWind sales records for January (rep, region, product, sales amount).

Dataset: [`../data/northwind-sales-jan.csv`](../data/northwind-sales-jan.csv)

| RepID | Rep Name | Region | Product | Sales Amount |
| --- | --- | --- | --- | --- |
| R001 | Aisha Rahman | Central | Office Chair | 12400 |
| R002 | Brian Tan | East | Standing Desk | 8600 |
| R003 | Chen Wei Ling | North | Filing Cabinet | 4750 |
| R004 | Devi Kumar | West | Office Chair | 15900 |
| R005 | Elson Goh | Central | Monitor Arm | 980 |
| R006 | Farah Ismail | East | Standing Desk | 22300 |
| R007 | Gopal Nair | North | Filing Cabinet | 3120 |
| R008 | Hui Ling Ong | West | Monitor Arm | 640 |
| R009 | Irfan Salleh | Central | Standing Desk | 9450 |
| R010 | Jasmine Lee | East | Office Chair | 5300 |
| R011 | Kelvin Yeo | North | Monitor Arm | 1780 |
| R012 | Lina Hassan | West | Filing Cabinet | 7250 |

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 1: Set Up the Developer Ribbon and Save a Macro-Enabled Workbook.

Business context: You have just joined NorthWind Trading as a sales analyst. Before you can automate anything, you must switch on Excel's Developer tools, open the Visual Basic Editor, and learn why a workbook containing macros has to be saved as .xlsm.
Workbook goal: A macro-enabled workbook, NorthWind-Commission.xlsm, with the Developer ribbon enabled and the VBA Editor open.
Available workbook context: Mock data: 12 NorthWind sales records for January (rep, region, product, sales amount).
Required VBA concepts/tools: Excel, Developer ribbon, Visual Basic Editor
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: The Developer tab is visible, Alt+F11 opens the Visual Basic Editor, the Sales sheet holds 12 data rows, and the title bar reads NorthWind-Commission.xlsm.

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
- [ ] Evidence: the code passes this acceptance test — The Developer tab is visible, Alt+F11 opens the Visual Basic Editor, the Sales sheet holds 12 data rows, and the title bar reads NorthWind-Commission.xlsm.

### Refine with evidence

```text
The VBA did not yet pass Lab 1. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: The Developer tab is visible, Alt+F11 opens the Visual Basic Editor, the Sales sheet holds 12 data rows, and the title bar reads NorthWind-Commission.xlsm.
```

## Step-by-step

1. Open Excel and create a new blank workbook.
2. Click File → Options → Customize the Ribbon, tick 'Developer' in the right-hand list, then click OK.
3. Confirm the Developer tab now appears on the ribbon, and click Developer → Visual Basic.
4. In the VBA Editor, note the Project Explorer on the left showing VBAProject (Book1) with its Sheet and ThisWorkbook objects.
5. Close the editor with Alt+Q, then import the mock data: paste the northwind-sales-jan.csv contents into Sheet1 starting at cell A1.
6. Rename Sheet1 to 'Sales' by double-clicking the sheet tab.
7. Click File → Save As, choose 'Excel Macro-Enabled Workbook (*.xlsm)', and save as NorthWind-Commission.xlsm.
8. Deliberately try File → Save As with the plain .xlsx type and read Excel's warning that macro-free workbooks discard VBA — then cancel.

## The code

```vba
Sub VerifyVBAEnvironment()
    ' AI-generated environment check for Lab 1
    MsgBox "VBA is enabled. Save this workbook as .xlsm before continuing.", _
           vbInformation, "AI Vibe Coding for Excel VBA"
End Sub
```

## Test it
The Developer tab is visible, Alt+F11 opens the Visual Basic Editor, the Sales sheet holds 12 data rows, and the title bar reads NorthWind-Commission.xlsm.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-01-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*