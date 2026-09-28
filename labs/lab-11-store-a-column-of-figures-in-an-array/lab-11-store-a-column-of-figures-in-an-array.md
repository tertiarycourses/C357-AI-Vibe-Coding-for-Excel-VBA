# Lab 11 — Store a Column of Figures in an Array

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 02:** VBA Programming  
**Learning outcome:** LO2

## Objective
This lab develops your ability to declare an array and assign values to it by index.

## Goal
Reading the same cells repeatedly is slow and fragile. You will load NorthWind's six monthly shipping charges into an array, then read them back by index — the foundation for loop processing.

## What you'll build
A ShippingArray Sub that loads six monthly charges into an array and reports them.

**Tools used:** Dim with parentheses, array index, Offset

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-11-start.xlsm`](lab-11-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-11-end.xlsm`](lab-11-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 10. If you already have your own workbook from Lab 10, keep using it — `lab-11-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Mock data: six months of shipping charges for the Central region.

Dataset: [`../data/northwind-shipping.csv`](../data/northwind-shipping.csv)

| Month | Shipping Charge |
| --- | --- |
| January | 320.5 |
| February | 415.75 |
| March | 298.0 |
| April | 502.25 |
| May | 388.4 |
| June | 447.9 |

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 11: Store a Column of Figures in an Array.

Business context: Reading the same cells repeatedly is slow and fragile. You will load NorthWind's six monthly shipping charges into an array, then read them back by index — the foundation for loop processing.
Workbook goal: A ShippingArray Sub that loads six monthly charges into an array and reports them.
Available workbook context: Mock data: six months of shipping charges for the Central region.
Required VBA concepts/tools: Dim with parentheses, array index, Offset
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: The message box lists all six charges beginning with 320.5 and ending with 447.9.

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
- [ ] Evidence: the code passes this acceptance test — The message box lists all six charges beginning with 320.5 and ending with 447.9.

### Refine with evidence

```text
The VBA did not yet pass Lab 11. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: The message box lists all six charges beginning with 320.5 and ending with 447.9.
```

## Step-by-step

1. Add a sheet named 'Shipping' and paste in the northwind-shipping.csv mock data.
2. Write the ShippingArray Sub and declare Dim curCharges(5) As Currency — six elements, indexed 0 to 5.
3. Declare Dim iCounter As Integer to hold the array index.
4. Activate the Shipping sheet and select cell B2, the first charge.
5. Assign each cell to its array element using ActiveCell.Offset(iCounter, 0).Value.
6. Read the array back and concatenate the six values into one string.
7. Display the string in a MsgBox and confirm all six charges appear in order.
8. Note in a comment why the array is declared as (5) yet holds six values — indexing starts at 0.

## The code

```vba
Sub ShippingArray()

    Dim curCharges(5) As Currency     ' 6 elements, indexed 0 to 5
    Dim iCounter As Integer
    Dim strOut As String

    Worksheets("Shipping").Activate
    Range("B2").Activate

    For iCounter = 0 To 5
        curCharges(iCounter) = ActiveCell.Offset(iCounter, 0).Value
    Next iCounter

    For iCounter = 0 To 5
        strOut = strOut & "Month " & (iCounter + 1) & ": " & _
                 curCharges(iCounter) & Chr(10)
    Next iCounter

    MsgBox strOut

End Sub
```

## Test it
The message box lists all six charges beginning with 320.5 and ending with 447.9.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-11-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*