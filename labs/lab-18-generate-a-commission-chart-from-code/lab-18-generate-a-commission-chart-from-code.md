# Lab 18 — Generate a Commission Chart from Code

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 03:** Function and Sub Procedure  
**Learning outcome:** LO3

## Objective
This lab develops your ability to produce a chart programmatically to visualise results.

## Goal
The sales manager wants the commission breakdown as a chart, refreshed on demand. You will write a macro that creates a bar chart from the commission range and attach it to a button — the same capability the capstone requires.

## What you'll build
A PlotCommission macro that builds a bar chart, wired to a 'Plot Commission' button.

**Tools used:** Shapes.AddChart, ChartType, xlColumnClustered, SetSourceData, Form Control button

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-18-start.xlsm`](lab-18-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-18-end.xlsm`](lab-18-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 17. If you already have your own workbook from Lab 17, keep using it — `lab-18-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Charts the rep names and commission values on the Commission sheet.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 18: Generate a Commission Chart from Code.

Business context: The sales manager wants the commission breakdown as a chart, refreshed on demand. You will write a macro that creates a bar chart from the commission range and attach it to a button — the same capability the capstone requires.
Workbook goal: A PlotCommission macro that builds a bar chart, wired to a 'Plot Commission' button.
Available workbook context: Charts the rep names and commission values on the Commission sheet.
Required VBA concepts/tools: Shapes.AddChart, ChartType, xlColumnClustered, SetSourceData, Form Control button
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Clicking the 'Plot Commission' button draws a clustered column chart of commission by rep.

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
- [ ] Evidence: the code passes this acceptance test — Clicking the 'Plot Commission' button draws a clustered column chart of commission by rep.

### Refine with evidence

```text
The VBA did not yet pass Lab 18. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Clicking the 'Plot Commission' button draws a clustered column chart of commission by rep.
```

## Step-by-step

1. Open the Commission sheet and confirm rep names sit in column B and commissions in column D.
2. Write a Sub named PlotCommission in Module1.
3. Add a chart with ActiveSheet.Shapes.AddChart.Select.
4. Set the chart type with ActiveChart.ChartType = xlColumnClustered.
5. Point the chart at the data using ActiveChart.SetSourceData Source:=Range("B1:B13,D1:D13").
6. Add a chart title naming it 'NorthWind Commission by Rep'.
7. Run the macro and confirm the bar chart appears on the sheet.
8. Insert a Form Control button labelled 'Plot Commission' and assign the PlotCommission macro to it.

## The code

```vba
Sub PlotCommission()

    ' Builds a clustered column chart of commission by sales rep
    Worksheets("Commission").Activate

    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range("B1:B13,D1:D13")
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "NorthWind Commission by Rep"

End Sub
```

## Test it
Clicking the 'Plot Commission' button draws a clustered column chart of commission by rep.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-18-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*