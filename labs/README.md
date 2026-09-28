# AI Vibe Coding for Excel VBA — Hands-On Labs

**Course Code:** C357  |  **Version v1.0 · 28 September 2026**

Every lab in this course builds one realistic deliverable: an automated commission tracker for the sales team of NorthWind Trading, a Singapore distributor. You start by framing a task for an AI coding assistant, and finish with a reviewed and tested UserForm solution that captures a sale, calculates banded commission, handles bad input and plots the result. Every lab retains its prompt, generated VBA and test evidence.

All 31 labs build one workbook, **NorthWind-Commission.xlsm**, from a blank file to a complete automated commission tracker. Work through them in order — each lab assumes the previous one.

## Before you start

- Microsoft Excel 2016 or later (or Microsoft 365) on Windows or Mac.
- The Developer ribbon enabled — Lab 1 walks you through it.
- The mock datasets in [`data/`](data/).
- Full instructions and worked solutions are also in the Learner Guide.

## Each lab ships its own Excel files

Every lab has its own folder containing a **start** and an **end** workbook, embedded runnable VBA, and the completed source ready to inspect or import:

```
labs/lab-NN-<title>/
    lab-NN-<title>.md      the lab instructions
    lab-NN-start.xlsm      open this one — the workbook before the lab
    lab-NN-end.xlsm        the workbook after the lab — sheets and results
    vba/                   importable .bas / .cls / .frm source
```

Because all 31 labs build the same workbook, **each lab's start file is the previous lab's end file**. Work straight through with one workbook of your own, or pick up the start file of any lab to rejoin the course at that point.

> **About the macros** — every `.xlsm` contains a real embedded VBA project. The start workbook contains a safe lab-launch macro; the end workbook contains a runnable lab-specific checkpoint. The complete worked source, including event and UserForm code that belongs in a specific component, is supplied in `vba/` for inspection, comparison and import. Always review AI-generated VBA before running it and save your work as **.xlsm**.

## Topic 01 — Get Started on VBA

*What VBA is · The Visual Basic Editor · Modules & procedures · Message and input boxes · Recording macros*

| Lab | Title | You'll build | Excel files |
| --- | --- | --- | --- |
| 1 | [Set Up the Developer Ribbon and Save a Macro-Enabled Workbook](lab-01-set-up-the-developer-ribbon-and-save-a-macro-enabled-workbook/lab-01-set-up-the-developer-ribbon-and-save-a-macro-enabled-workbook.md) | A macro-enabled workbook, NorthWind-Commission.xlsm, with the Developer ribbon enabled and the VBA Editor open. | [start](lab-01-set-up-the-developer-ribbon-and-save-a-macro-enabled-workbook/lab-01-start.xlsm) · [end](lab-01-set-up-the-developer-ribbon-and-save-a-macro-enabled-workbook/lab-01-end.xlsm) |
| 2 | [Write Your First Macro in a Module](lab-02-write-your-first-macro-in-a-module/lab-02-write-your-first-macro-in-a-module.md) | A Module1 containing a Welcome Sub procedure that greets the user with a message box. | [start](lab-02-write-your-first-macro-in-a-module/lab-02-start.xlsm) · [end](lab-02-write-your-first-macro-in-a-module/lab-02-end.xlsm) |
| 3 | [Message Boxes and Button Choices](lab-03-message-boxes-and-button-choices/lab-03-message-boxes-and-button-choices.md) | A ConfirmRecalculate Sub that asks the user to proceed and reports the button they chose. | [start](lab-03-message-boxes-and-button-choices/lab-03-start.xlsm) · [end](lab-03-message-boxes-and-button-choices/lab-03-end.xlsm) |
| 4 | [Capture User Input with an Input Box](lab-04-capture-user-input-with-an-input-box/lab-04-capture-user-input-with-an-input-box.md) | A ShowCommission Sub that asks for a commission rate and reports the commission on the selected cell. | [start](lab-04-capture-user-input-with-an-input-box/lab-04-start.xlsm) · [end](lab-04-capture-user-input-with-an-input-box/lab-04-end.xlsm) |
| 5 | [Record a Macro to Calculate Sales Tax](lab-05-record-a-macro-to-calculate-sales-tax/lab-05-record-a-macro-to-calculate-sales-tax.md) | A recorded Cal_SalesTax macro plus your annotated reading of the generated code. | [start](lab-05-record-a-macro-to-calculate-sales-tax/lab-05-start.xlsm) · [end](lab-05-record-a-macro-to-calculate-sales-tax/lab-05-end.xlsm) |
| 6 | [Run a Macro from a Button and a Shape](lab-06-run-a-macro-from-a-button-and-a-shape/lab-06-run-a-macro-from-a-button-and-a-shape.md) | A Sales sheet with a 'Calculate Commission' button and a shape that both launch macros. | [start](lab-06-run-a-macro-from-a-button-and-a-shape/lab-06-start.xlsm) · [end](lab-06-run-a-macro-from-a-button-and-a-shape/lab-06-end.xlsm) |

## Topic 02 — VBA Programming

*Variables & constants · Data types & arrays · Operators · Decisions · Loops*

| Lab | Title | You'll build | Excel files |
| --- | --- | --- | --- |
| 7 | [Declare Variables and Constants with the Right Data Type](lab-07-declare-variables-and-constants-with-the-right-data-type/lab-07-declare-variables-and-constants-with-the-right-data-type.md) | A CalculateNetCommission Sub using Dim variables with correct types and a Const GST rate. | [start](lab-07-declare-variables-and-constants-with-the-right-data-type/lab-07-start.xlsm) · [end](lab-07-declare-variables-and-constants-with-the-right-data-type/lab-07-end.xlsm) |
| 8 | [Build Expressions with VBA Operators](lab-08-build-expressions-with-vba-operators/lab-08-build-expressions-with-vba-operators.md) | A SummariseRep Sub that computes figures and concatenates a one-line summary sentence. | [start](lab-08-build-expressions-with-vba-operators/lab-08-start.xlsm) · [end](lab-08-build-expressions-with-vba-operators/lab-08-end.xlsm) |
| 9 | [Branch with If…Then…ElseIf…Else](lab-09-branch-with-ifthenelseifelse/lab-09-branch-with-ifthenelseifelse.md) | An ElseIfCommission Sub that applies the correct banded rate to the selected sale. | [start](lab-09-branch-with-ifthenelseifelse/lab-09-start.xlsm) · [end](lab-09-branch-with-ifthenelseifelse/lab-09-end.xlsm) |
| 10 | [Simplify Banded Logic with Select Case](lab-10-simplify-banded-logic-with-select-case/lab-10-simplify-banded-logic-with-select-case.md) | A SelectRate Sub implementing four commission bands with Select Case. | [start](lab-10-simplify-banded-logic-with-select-case/lab-10-start.xlsm) · [end](lab-10-simplify-banded-logic-with-select-case/lab-10-end.xlsm) |
| 11 | [Store a Column of Figures in an Array](lab-11-store-a-column-of-figures-in-an-array/lab-11-store-a-column-of-figures-in-an-array.md) | A ShippingArray Sub that loads six monthly charges into an array and reports them. | [start](lab-11-store-a-column-of-figures-in-an-array/lab-11-start.xlsm) · [end](lab-11-store-a-column-of-figures-in-an-array/lab-11-end.xlsm) |
| 12 | [Process Every Row with For…Next and For Each](lab-12-process-every-row-with-fornext-and-for-each/lab-12-process-every-row-with-fornext-and-for-each.md) | A CommissionAllRows Sub and a ListRegions Sub demonstrating both loop styles. | [start](lab-12-process-every-row-with-fornext-and-for-each/lab-12-start.xlsm) · [end](lab-12-process-every-row-with-fornext-and-for-each/lab-12-end.xlsm) |
| 13 | [Find the First Empty Row with Do…While and Do…Until](lab-13-find-the-first-empty-row-with-dowhile-and-dountil/lab-13-find-the-first-empty-row-with-dowhile-and-dountil.md) | A FindFirstEmptyRow Sub that locates and reports the first blank row in the data. | [start](lab-13-find-the-first-empty-row-with-dowhile-and-dountil/lab-13-start.xlsm) · [end](lab-13-find-the-first-empty-row-with-dowhile-and-dountil/lab-13-end.xlsm) |

## Topic 03 — Function and Sub Procedure

*Function vs Sub · Arguments & return values · Scope · Worksheet functions · Charts*

| Lab | Title | You'll build | Excel files |
| --- | --- | --- | --- |
| 14 | [Write a Reusable Function that Returns a Value](lab-14-write-a-reusable-function-that-returns-a-value/lab-14-write-a-reusable-function-that-returns-a-value.md) | A Cal_Commission function callable from both VBA and any worksheet cell. | [start](lab-14-write-a-reusable-function-that-returns-a-value/lab-14-start.xlsm) · [end](lab-14-write-a-reusable-function-that-returns-a-value/lab-14-end.xlsm) |
| 15 | [Compare a Function with a Sub Procedure](lab-15-compare-a-function-with-a-sub-procedure/lab-15-compare-a-function-with-a-sub-procedure.md) | A matched pair, Cal_Commission (Function) and ShowCommissionSub (Sub), plus a written comparison. | [start](lab-15-compare-a-function-with-a-sub-procedure/lab-15-start.xlsm) · [end](lab-15-compare-a-function-with-a-sub-procedure/lab-15-end.xlsm) |
| 16 | [Control Visibility with Public and Private Scope](lab-16-control-visibility-with-public-and-private-scope/lab-16-control-visibility-with-public-and-private-scope.md) | A module where internal helpers are Private and only the master macro is Public. | [start](lab-16-control-visibility-with-public-and-private-scope/lab-16-start.xlsm) · [end](lab-16-control-visibility-with-public-and-private-scope/lab-16-end.xlsm) |
| 17 | [Reuse Excel's Built-in Worksheet Functions](lab-17-reuse-excels-built-in-worksheet-functions/lab-17-reuse-excels-built-in-worksheet-functions.md) | A SalesStatistics Sub reporting total, average, highest and lowest sales, plus a PMT calculation. | [start](lab-17-reuse-excels-built-in-worksheet-functions/lab-17-start.xlsm) · [end](lab-17-reuse-excels-built-in-worksheet-functions/lab-17-end.xlsm) |
| 18 | [Generate a Commission Chart from Code](lab-18-generate-a-commission-chart-from-code/lab-18-generate-a-commission-chart-from-code.md) | A PlotCommission macro that builds a bar chart, wired to a 'Plot Commission' button. | [start](lab-18-generate-a-commission-chart-from-code/lab-18-start.xlsm) · [end](lab-18-generate-a-commission-chart-from-code/lab-18-end.xlsm) |

## Topic 04 — Managing Excel Objects

*Application, Workbook, Worksheet & Range objects · Errors · Error handling · Testing & debugging*

| Lab | Title | You'll build | Excel files |
| --- | --- | --- | --- |
| 19 | [Read and Write Cells with the Range Object](lab-19-read-and-write-cells-with-the-range-object/lab-19-read-and-write-cells-with-the-range-object.md) | A FormatSalesReport Sub that renames headings and writes calculated values into the sheet. | [start](lab-19-read-and-write-cells-with-the-range-object/lab-19-start.xlsm) · [end](lab-19-read-and-write-cells-with-the-range-object/lab-19-end.xlsm) |
| 20 | [Copy Data Between Worksheets](lab-20-copy-data-between-worksheets/lab-20-copy-data-between-worksheets.md) | A CopyToSummary Sub that consolidates the January data onto a Summary sheet. | [start](lab-20-copy-data-between-worksheets/lab-20-start.xlsm) · [end](lab-20-copy-data-between-worksheets/lab-20-end.xlsm) |
| 21 | [Automate Workbooks and Worksheets](lab-21-automate-workbooks-and-worksheets/lab-21-automate-workbooks-and-worksheets.md) | A MonthEndSetup Sub that adds and names a sheet, and reports on open workbooks. | [start](lab-21-automate-workbooks-and-worksheets/lab-21-start.xlsm) · [end](lab-21-automate-workbooks-and-worksheets/lab-21-end.xlsm) |
| 22 | [Identify Syntax, Runtime and Logical Errors](lab-22-identify-syntax-runtime-and-logical-errors/lab-22-identify-syntax-runtime-and-logical-errors.md) | Three annotated broken procedures demonstrating each error type and its symptom. | [start](lab-22-identify-syntax-runtime-and-logical-errors/lab-22-start.xlsm) · [end](lab-22-identify-syntax-runtime-and-logical-errors/lab-22-end.xlsm) |
| 23 | [Handle Runtime Errors with On Error GoTo](lab-23-handle-runtime-errors-with-on-error-goto/lab-23-handle-runtime-errors-with-on-error-goto.md) | A CheckError Sub with a labelled error handler that traps division-by-zero and bad input. | [start](lab-23-handle-runtime-errors-with-on-error-goto/lab-23-start.xlsm) · [end](lab-23-handle-runtime-errors-with-on-error-goto/lab-23-end.xlsm) |
| 24 | [Debug with Breakpoints, Step Into and the Watch Window](lab-24-debug-with-breakpoints-step-into-and-the-watch-window/lab-24-debug-with-breakpoints-step-into-and-the-watch-window.md) | A debugging session on the commission loop with a breakpoint and a watch expression. | [start](lab-24-debug-with-breakpoints-step-into-and-the-watch-window/lab-24-start.xlsm) · [end](lab-24-debug-with-breakpoints-step-into-and-the-watch-window/lab-24-end.xlsm) |
| 25 | [Find a Value and Handle the Not-Found Case](lab-25-find-a-value-and-handle-the-not-found-case/lab-25-find-a-value-and-handle-the-not-found-case.md) | A FindThisValue Sub that locates a rep and reports clearly when the name is not found. | [start](lab-25-find-a-value-and-handle-the-not-found-case/lab-25-start.xlsm) · [end](lab-25-find-a-value-and-handle-the-not-found-case/lab-25-end.xlsm) |

## Topic 05 — Create User Defined Forms

*Workbook & worksheet events · UserForms & controls · Event procedures · Documentation*

| Lab | Title | You'll build | Excel files |
| --- | --- | --- | --- |
| 26 | [Respond to Workbook Events](lab-26-respond-to-workbook-events/lab-26-respond-to-workbook-events.md) | ThisWorkbook event procedures that run on open and guard the close. | [start](lab-26-respond-to-workbook-events/lab-26-start.xlsm) · [end](lab-26-respond-to-workbook-events/lab-26-end.xlsm) |
| 27 | [Respond to Worksheet Events](lab-27-respond-to-worksheet-events/lab-27-respond-to-worksheet-events.md) | A Worksheet_Change event on the Commission sheet that confirms every manual edit. | [start](lab-27-respond-to-worksheet-events/lab-27-start.xlsm) · [end](lab-27-respond-to-worksheet-events/lab-27-end.xlsm) |
| 28 | [Design a UserForm for Sales Entry](lab-28-design-a-userform-for-sales-entry/lab-28-design-a-userform-for-sales-entry.md) | A frmSalesEntry UserForm laid out with named, captioned controls. | [start](lab-28-design-a-userform-for-sales-entry/lab-28-start.xlsm) · [end](lab-28-design-a-userform-for-sales-entry/lab-28-end.xlsm) |
| 29 | [Write the Form's Event Procedures](lab-29-write-the-forms-event-procedures/lab-29-write-the-forms-event-procedures.md) | A working frmSalesEntry that appends each submitted record to the Sales sheet. | [start](lab-29-write-the-forms-event-procedures/lab-29-start.xlsm) · [end](lab-29-write-the-forms-event-procedures/lab-29-end.xlsm) |
| 30 | [Validate Form Input and Document the Solution](lab-30-validate-form-input-and-document-the-solution/lab-30-validate-form-input-and-document-the-solution.md) | A validated frmSalesEntry plus an in-workbook ReadMe documenting the solution. | [start](lab-30-validate-form-input-and-document-the-solution/lab-30-start.xlsm) · [end](lab-30-validate-form-input-and-document-the-solution/lab-30-end.xlsm) |
| 31 | [Capstone — Complete the NorthWind Commission Tracker](lab-31-capstone-complete-the-northwind-commission-tracker/lab-31-capstone-complete-the-northwind-commission-tracker.md) | The completed NorthWind-Commission.xlsm automating capture, calculation, validation and reporting. | [start](lab-31-capstone-complete-the-northwind-commission-tracker/lab-31-start.xlsm) · [end](lab-31-capstone-complete-the-northwind-commission-tracker/lab-31-end.xlsm) |

## Mock datasets

| Dataset | Used by |
| --- | --- |
| [`data/northwind-sales-jan.csv`](data/northwind-sales-jan.csv) | Lab 1 — Set Up the Developer Ribbon and Save a Macro-Enabled Workbook |
| [`data/northwind-sales-tax.csv`](data/northwind-sales-tax.csv) | Lab 5 — Record a Macro to Calculate Sales Tax |
| [`data/northwind-targets.csv`](data/northwind-targets.csv) | Lab 8 — Build Expressions with VBA Operators |
| [`data/northwind-shipping.csv`](data/northwind-shipping.csv) | Lab 11 — Store a Column of Figures in an Array |
| [`data/northwind-loan.csv`](data/northwind-loan.csv) | Lab 17 — Reuse Excel's Built-in Worksheet Functions |

---

## Attend this course

These labs are run over 2 days as the course **AI Vibe Coding for Excel VBA** (C357). Register for an upcoming class:

### 👉 [www.tertiarycourses.com.sg/ai-vibe-coding-for-excel-vba.html](https://www.tertiarycourses.com.sg/ai-vibe-coding-for-excel-vba.html)

---

*© 2026 Tertiary Infotech Academy Pte Ltd (UEN 201200696W). All rights reserved.*