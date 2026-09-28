# Lab 12 — Process Every Row with For…Next and For Each

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to repeat operations over a range using counted and collection loops.

## Goal
Twelve reps today, two hundred next year — the tool must scale. You will calculate commission for every row with a For…Next loop, then use For Each to walk a list of sales regions, and reverse the loop's direction to see how Step works.

## What you'll build
A CommissionAllRows Sub and a ListRegions Sub demonstrating both loop styles.

**Tools used:** For…Next, Step, For Each…Next, Exit For, Cells

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-12-start.xlsm`](lab-12-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-12-end.xlsm`](lab-12-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 11. If you already have your own workbook from Lab 11, keep using it — `lab-12-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Commission sheet from Lab 10 (12 rows) plus a regions array.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 12: Process Every Row with For…Next and For Each.

Business context: Twelve reps today, two hundred next year — the tool must scale. You will calculate commission for every row with a For…Next loop, then use For Each to walk a list of sales regions, and reverse the loop's direction to see how Step works.
Workbook goal: A CommissionAllRows Sub and a ListRegions Sub demonstrating both loop styles.
Available workbook context: Uses the Commission sheet from Lab 10 (12 rows) plus a regions array.
Required VBA concepts/tools: For…Next, Step, For Each…Next, Exit For, Cells
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: All 12 commission cells populate in one run, and ListRegions stops after reporting North.

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
- [ ] Evidence: the code passes this acceptance test — All 12 commission cells populate in one run, and ListRegions stops after reporting North.

### Refine with evidence

```text
The VBA did not yet pass Lab 12. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: All 12 commission cells populate in one run, and ListRegions stops after reporting North.
```

## Step-by-step

1. Open the Commission sheet holding the 12 NorthWind rows.
2. Write the CommissionAllRows Sub declaring iRow As Integer.
3. Loop with For iRow = 2 To 13 so the header row is skipped.
4. Inside the loop, read Cells(iRow, 3).Value and write the 10% commission into Cells(iRow, 4).
5. Run the macro and confirm all 12 commissions are filled in one click.
6. Write a second loop using Step -1 to walk from row 13 back to row 2, and observe the reversed order.
7. Write the ListRegions Sub with an array of Central, East, North and West, and walk it with For Each.
8. Add an Exit For that stops the loop as soon as 'North' is reached, and confirm the early exit.

## The code

```vba
Sub CommissionAllRows()

    Dim iRow As Integer

    Worksheets("Commission").Activate

    For iRow = 2 To 13                       ' skip the header row
        Cells(iRow, 4).Value = Cells(iRow, 3).Value * 0.1
    Next iRow

    MsgBox "All commissions calculated."

End Sub


Sub CommissionReversed()

    Dim iRow As Integer

    For iRow = 13 To 2 Step -1               ' Step -1 walks backwards
        Debug.Print Cells(iRow, 2).Value
    Next iRow

End Sub


Sub ListRegions()

    Dim varRegion As Variant
    Dim strRegions As Variant

    strRegions = Array("Central", "East", "North", "West")

    For Each varRegion In strRegions
        MsgBox varRegion
        If varRegion = "North" Then Exit For      ' early exit
    Next

End Sub
```

## Test it
All 12 commission cells populate in one run, and ListRegions stops after reporting North.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-12-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*