# Lab 24 — Debug with Breakpoints, Step Into and the Watch Window

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to test and debug code using the VBA Editor's debugging tools.

## Goal
A commission figure is wrong and you cannot see why by reading the code. You will pause the macro on a breakpoint, walk it one line at a time, and watch a variable change value as the loop runs.

## What you'll build
A debugging session on the commission loop with a breakpoint and a watch expression.

**Tools used:** Breakpoints (F9), Step Into (F8), Watch window, Immediate window, Debug.Print

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-24-start.xlsm`](lab-24-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-24-end.xlsm`](lab-24-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 23. If you already have your own workbook from Lab 23, keep using it — `lab-24-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the CommissionAllRows loop from Lab 12.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 24: Debug with Breakpoints, Step Into and the Watch Window.

Business context: A commission figure is wrong and you cannot see why by reading the code. You will pause the macro on a breakpoint, walk it one line at a time, and watch a variable change value as the loop runs.
Workbook goal: A debugging session on the commission loop with a breakpoint and a watch expression.
Available workbook context: Uses the CommissionAllRows loop from Lab 12.
Required VBA concepts/tools: Breakpoints (F9), Step Into (F8), Watch window, Immediate window, Debug.Print
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Execution pauses at the breakpoint, F8 advances one line per press, and the Immediate window lists a commission line per row.

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
- [ ] Evidence: the code passes this acceptance test — Execution pauses at the breakpoint, F8 advances one line per press, and the Immediate window lists a commission line per row.

### Refine with evidence

```text
The VBA did not yet pass Lab 24. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Execution pauses at the breakpoint, F8 advances one line per press, and the Immediate window lists a commission line per row.
```

## Step-by-step

1. Open the CommissionAllRows procedure in the VBA Editor.
2. Click in the left margin beside the line inside the loop to set a breakpoint — a red dot appears.
3. Run the macro; execution pauses at the breakpoint with the line highlighted yellow.
4. Press F8 to Step Into and walk forward one line at a time, watching the highlight move.
5. Hover the mouse over the iRow variable to see its current value in a tooltip.
6. Right-click the commission variable and choose Add Watch to track it in the Watch window.
7. Press F8 repeatedly through several loop iterations and observe the watched value changing.
8. Add Debug.Print to write each row's commission to the Immediate window (Ctrl+G), then run to completion.

## The code

```vba
Sub CommissionAllRowsDebug()

    Dim iRow As Integer
    Dim curCommission As Currency

    Worksheets("Commission").Activate

    For iRow = 2 To 13
        ' <- set a breakpoint on the next line (F9)
        curCommission = Cells(iRow, 3).Value * 0.1
        Cells(iRow, 4).Value = curCommission

        ' Ctrl+G opens the Immediate window to read this output
        Debug.Print "Row " & iRow & ": " & _
                    Cells(iRow, 2).Value & " = " & curCommission
    Next iRow

End Sub
```

## Test it
Execution pauses at the breakpoint, F8 advances one line per press, and the Immediate window lists a commission line per row.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-24-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*