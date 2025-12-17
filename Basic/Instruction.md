# Excel VBA Basic Training: Printing Company Dataset and Exercises

This Markdown file contains the sample dummy dataset for a printing company (in table format for easy copying into Excel), and step-by-step instructions for basic MS Excel VBA training.

## Sample Dummy Dataset for Printing Company

Copy the table below and paste it into a new Excel sheet (select the entire table, copy, then paste into cell A1 in Excel).  
The "Total Price" column uses formulas (`=Quantity * Price per Unit`). Enter them manually as shown.

| Order ID | Customer Name    | Product Type   | Quantity | Price per Unit | Total Price       | Order Date  | Status    |
|----------|------------------|----------------|----------|----------------|-------------------|-------------|-----------|
| 1        | John Doe        | Brochures     | 500     | 0.25          | =D2*E2           | 2025-01-15 | Pending   |
| 2        | Jane Smith      | Flyers        | 1000    | 0.15          | =D3*E3           | 2025-02-20 | Completed |
| 3        | Acme Corp       | Business Cards| 2000    | 0.10          | =D4*E4           | 2025-03-05 | Pending   |
| 4        | Bob Johnson     | Posters       | 100     | 2.50          | =D5*E5           | 2025-04-10 | Completed |
| 5        | Alice Wonderland| Banners       | 50      | 5.00          | =D6*E6           | 2025-05-15 | Pending   |
| 6        | Tech Inc        | Brochures     | 750     | 0.20          | =D7*E7           | 2025-06-20 | Completed |
| 7        | Foodies Ltd     | Menus         | 300     | 0.40          | =D8*E8           | 2025-07-25 | Pending   |
| 8        | Sports Gear     | Flyers        | 1500    | 0.12          | =D9*E9           | 2025-08-30 | Completed |
| 9        | Book Club       | Booklets      | 200     | 1.00          | =D10*E10         | 2025-09-05 | Pending   |
| 10       | Event Planners  | Invitations   | 400     | 0.30          | =D11*E11         | 2025-10-10 | Completed |

**Notes:**
- Columns: A = Order ID, B = Customer Name, C = Product Type, D = Quantity, E = Price per Unit, F = Total Price, G = Order Date, H = Status.
- After pasting, drag the Total Price formula down if needed.
- Save the file as an Excel Macro-Enabled Workbook (`.xlsm`) to enable VBA macros.

## Step-by-Step Instructions for Basic MS Excel VBA Training

This training focuses on basic VBA concepts using the dataset. We'll cover opening the VBA editor, writing simple macros (Subs), using loops, variables, cell manipulation, and conditional logic. Assume the data is in Sheet1, starting at A1.

### Prerequisites
- Open the dataset in Microsoft Excel.
- Enable macros if prompted.
- For VBA to run, save the file as `.xlsm` (Excel Macro-Enabled Workbook): File > Save As > Excel Macro-Enabled Workbook.
- Enable the Developer tab: File > Options > Customize Ribbon > check "Developer".

### Step 1: Open the VBA Editor
1. In Excel, press `Alt + F11` (or Developer tab > Visual Basic).
2. In the VBA window, Project Explorer (Ctrl + R if not visible).
3. To insert a module: Insert > Module (creates Module1 for code).

### Step 2: Basic Macro - Display a Message Box

```vba
Sub HelloWorld()
    MsgBox "Welcome to Printing Company VBA Training!"
End Sub
```

**How to run:**
- Paste into Module1.
- Close VBA editor (Alt + Q).
- Run: Developer > Macros > HelloWorld > Run (or Alt + F8).

### Step 3: Calculate Total Prices (If Formulas Are Removed)

Clear column F first for practice.

```vba
Sub CalculateTotals()
    Dim i As Integer
    For i = 2 To 11
        Cells(i, 6).Value = Cells(i, 4).Value * Cells(i, 5).Value
    Next i
    MsgBox "Totals calculated!"
End Sub
```

**Explanation:** Loops through rows 2-11, calculates Total Price.

### Step 4: Conditional Macro - Update Status Based on Date

Sets Status to "Overdue" if Order Date < today.

```vba
Sub UpdateStatus()
    Dim i As Integer
    Dim today As Date
    today = Date
    
    For i = 2 To 11
        If Cells(i, 7).Value < today Then
            Cells(i, 8).Value = "Overdue"
        Else
            Cells(i, 8).Value = "On Time"
        End If
    Next i
    MsgBox "Statuses updated!"
End Sub
```

### Step 5: Formatting Macro - Highlight High-Quantity Orders

Highlights rows where Quantity > 500.

```vba
Sub HighlightHighQuantity()
    Dim i As Integer
    For i = 2 To 11
        If Cells(i, 4).Value > 500 Then
            Range("A" & i & ":H" & i).Interior.Color = vbYellow
        End If
    Next i
    MsgBox "High-quantity orders highlighted!"
End Sub
```

### Step 6: Sum All Totals and Display

```vba
Sub GrandTotal()
    Dim total As Double
    Dim i As Integer
    total = 0
    For i = 2 To 11
        total = total + Cells(i, 6).Value
    Next i
    MsgBox "Grand Total: " & Format(total, "Currency")
End Sub
```

### Practice Tips
- Debug: F8 to step through code.
- Use Macro Recorder to learn from actions.
- Expand: Add InputBox for user filters.

This covers basic VBA: Subs, variables, loops, conditions, cell access, and formatting. Enjoy practicing!
```

To download:
1. Copy everything above (from `# Excel VBA Basic Training...` to the end).
2. Paste into a text editor.
3. Save as `Printing_Company_VBA_Training.md`. 

You now have the complete content in one clean Markdown file!