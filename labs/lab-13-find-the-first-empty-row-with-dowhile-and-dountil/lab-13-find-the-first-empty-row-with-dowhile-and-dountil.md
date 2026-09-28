# Lab 13 — Find the First Empty Row with Do…While and Do…Until

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to repeat operations while or until a condition is met.

## Goal
New sales records must be appended below the existing ones, so the macro has to find the first empty row. You will solve this with a Do loop — the same technique your UserForm will use in Topic 5 to append records.

## What you'll build
A FindFirstEmptyRow Sub that locates and reports the first blank row in the data.

**Tools used:** Do While, Do Until, Loop, Exit Do, Offset, Activate

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-13-start.xlsm`](lab-13-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-13-end.xlsm`](lab-13-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 12. If you already have your own workbook from Lab 12, keep using it — `lab-13-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet; the mock data ends at row 13 so row 14 is the expected result.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 13: Find the First Empty Row with Do…While and Do…Until.

Business context: New sales records must be appended below the existing ones, so the macro has to find the first empty row. You will solve this with a Do loop — the same technique your UserForm will use in Topic 5 to append records.
Workbook goal: A FindFirstEmptyRow Sub that locates and reports the first blank row in the data.
Available workbook context: Uses the Sales sheet; the mock data ends at row 13 so row 14 is the expected result.
Required VBA concepts/tools: Do While, Do Until, Loop, Exit Do, Offset, Activate
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Both the Do Until and Do While versions report 'First empty row: 14'.

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
- [ ] Evidence: the code passes this acceptance test — Both the Do Until and Do While versions report 'First empty row: 14'.

### Refine with evidence

```text
The VBA did not yet pass Lab 13. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Both the Do Until and Do While versions report 'First empty row: 14'.
```

## Step-by-step

1. Open the Sales sheet and confirm the data occupies rows 2 to 13.
2. Write the FindFirstEmptyRow Sub and activate cell A2 as the starting point.
3. Open a loop with Do Until ActiveCell.Value = "".
4. Inside the loop, move down one row with ActiveCell.Offset(1, 0).Activate.
5. Close the loop with the Loop keyword and report the found row with ActiveCell.Row.
6. Run the macro and confirm it reports row 14.
7. Rewrite the same logic as Do While ActiveCell.Value <> "" and confirm the identical result.
8. Add a safety counter that uses Exit Do after 1000 iterations so a bad condition cannot hang Excel.

## The code

```vba
Sub FindFirstEmptyRow()

    Dim iSafety As Long

    Worksheets("Sales").Activate
    Range("A2").Activate

    Do Until ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Activate
        iSafety = iSafety + 1
        If iSafety > 1000 Then Exit Do        ' never hang Excel
    Loop

    MsgBox "First empty row: " & ActiveCell.Row

End Sub


Sub FindFirstEmptyRowWhile()

    Worksheets("Sales").Activate
    Range("A2").Activate

    Do While ActiveCell.Value <> ""           ' same result, opposite test
        ActiveCell.Offset(1, 0).Activate
    Loop

    MsgBox "First empty row: " & ActiveCell.Row

End Sub
```

## Test it
Both the Do Until and Do While versions report 'First empty row: 14'.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-13-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*