# Lab 19 — Read and Write Cells with the Range Object

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to manage worksheet data through the Range object.

## Goal
Everything a macro does to data goes through the Range object. You will read values, write values, rename a column heading and move around the sheet with Offset — the core skills for automating NorthWind's monthly report.

## What you'll build
A FormatSalesReport Sub that renames headings and writes calculated values into the sheet.

**Tools used:** Range, .Value, .Offset, .Activate, Cells

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-19-start.xlsm`](lab-19-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-19-end.xlsm`](lab-19-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 18. If you already have your own workbook from Lab 18, keep using it — `lab-19-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet mock data from Lab 1.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 19: Read and Write Cells with the Range Object.

Business context: Everything a macro does to data goes through the Range object. You will read values, write values, rename a column heading and move around the sheet with Offset — the core skills for automating NorthWind's monthly report.
Workbook goal: A FormatSalesReport Sub that renames headings and writes calculated values into the sheet.
Available workbook context: Uses the Sales sheet mock data from Lab 1.
Required VBA concepts/tools: Range, .Value, .Offset, .Activate, Cells
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Cell F1 reads 'Commission' and F2 holds the calculated commission for the first rep.

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
- [ ] Evidence: the code passes this acceptance test — Cell F1 reads 'Commission' and F2 holds the calculated commission for the first rep.

### Refine with evidence

```text
The VBA did not yet pass Lab 19. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Cell F1 reads 'Commission' and F2 holds the calculated commission for the first rep.
```

## Step-by-step

1. Open the Sales sheet and note that the commission column heading is still generic.
2. Write a Sub named FormatSalesReport in Module1.
3. Set the heading with Worksheets("Sales").Range("F1").Value = "Commission".
4. Read a single sale with Range("E2").Value and store it in a Currency variable.
5. Write the calculated commission back with Range("F2").Value.
6. Activate cell A1 and use ActiveCell.Offset(1, 4).Activate to step to the first sales figure.
7. Use Cells(2, 6) to write to the same commission cell by row and column number, and compare the two styles.
8. Run the macro and confirm the heading and the value both appear.

## The code

```vba
Sub FormatSalesReport()

    Dim curSale As Currency

    ' Range(): address you can see on the sheet
    Worksheets("Sales").Range("F1").Value = "Commission"

    curSale = Worksheets("Sales").Range("E2").Value
    Worksheets("Sales").Range("F2").Value = curSale * 0.1

    ' Offset(): move relative to the active cell
    Range("A1").Activate
    ActiveCell.Offset(1, 4).Activate        ' lands on E2

    ' Cells(row, column): address by number - ideal inside loops
    Cells(2, 6).Value = curSale * 0.1

End Sub
```

## Test it
Cell F1 reads 'Commission' and F2 holds the calculated commission for the first rep.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-19-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*