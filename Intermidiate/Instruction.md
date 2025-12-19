
# Excel VBA Intermediate Training: Printing Company Dataset and Exercises

This Markdown file contains the same sample dummy dataset for a printing company (in table format for easy copying into Excel), and **intermediate-level** MS Excel VBA exercises. These build on basic concepts by introducing:

- Working with Ranges and Worksheets
- UserForms (simple input form)
- Arrays and Collections
- Error Handling
- Events (Worksheet_Change)
- Functions (custom UDF)
- Working with multiple sheets and data validation

## Sample Dummy Dataset for Printing Company

Copy the table below and paste it into a new Excel sheet named **Orders** (select the table, copy, paste into cell A1).

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

**Setup Instructions:**
1. Create a new workbook, save as `.xlsm`.
2. Rename Sheet1 to **Orders** and paste the data.
3. Add two more sheets: **Summary** and **Products**.
4. In **Products** sheet (A1:B6), add a price list:

| Product Type    | Standard Price |
|-----------------|----------------|
| Brochures      | 0.25          |
| Flyers         | 0.15          |
| Business Cards | 0.10          |
| Posters        | 2.50          |
| Banners        | 5.00          |
| Others         | 0.50          |

## Intermediate VBA Exercises

All standard module code goes in a Module (Insert > Module). Event code goes in the worksheet/module specified.

### Exercise 1: Custom Function (UDF) - Lookup Price from Products Sheet

Create a function to auto-fill Price per Unit based on Product Type.

```vba
Function GetPrice(product As String) As Double
    Dim ws As Worksheet
    Set ws = Worksheets("Products")
    
    Dim lookupRange As Range
    Set lookupRange = ws.Range("A:B")
    
    Dim foundCell As Range
    Set foundCell = lookupRange.Find(What:=product, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        GetPrice = foundCell.Offset(0, 1).Value
    Else
        GetPrice = 0.5  ' Default for "Others"
    End If
End Function
```

**Usage:** In Orders sheet, column E (Price per Unit), enter `=GetPrice(C2)` and drag down.

### Exercise 2: Macro with Error Handling - Add New Order Safely

```vba
Sub AddNewOrder()
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = Worksheets("Orders").Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    Dim qty As Integer, price As Double, prod As String
    
    prod = InputBox("Enter Product Type:")
    qty = CInt(InputBox("Enter Quantity:"))
    price = GetPrice(prod)  ' Uses the UDF above
    
    With Worksheets("Orders")
        .Cells(lastRow, 1).Value = lastRow - 1  ' Order ID
        .Cells(lastRow, 2).Value = InputBox("Enter Customer Name:")
        .Cells(lastRow, 3).Value = prod
        .Cells(lastRow, 4).Value = qty
        .Cells(lastRow, 5).Value = price
        .Cells(lastRow, 6).Value = qty * price
        .Cells(lastRow, 7).Value = Date
        .Cells(lastRow, 8).Value = "Pending"
    End With
    
    MsgBox "New order added successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & ". Operation cancelled."
End Sub
```

### Exercise 3: Use Arrays for Faster Processing - Calculate All Totals

```vba
Sub CalculateAllTotalsFast()
    Dim dataRange As Range
    Set dataRange = Worksheets("Orders").Range("D2:H11")  ' Quantity to Status
    
    Dim dataArray As Variant
    dataArray = dataRange.Value  ' Load into array (faster than cell-by-cell)
    
    Dim i As Long
    For i = 1 To UBound(dataArray, 1)
        dataArray(i, 3) = dataArray(i, 1) * dataArray(i, 2)  ' Total = Qty * Price (column 3 in array)
    Next i
    
    dataRange.Value = dataArray  ' Write back in one operation
    
    MsgBox "All totals recalculated using arrays!"
End Sub
```

### Exercise 4: Worksheet Event - Auto-Update Status on Date Change

Paste this code in the **Orders** worksheet module (right-click sheet tab > View Code).

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("G:G")) Is Nothing Then  ' If Order Date changed
        Application.EnableEvents = False
        
        Dim cell As Range
        For Each cell In Target
            If cell.Value < Date Then
                cell.Offset(0, 1).Value = "Overdue"
            Else
                cell.Offset(0, 1).Value = "On Time"
            End If
        Next cell
        
        Application.EnableEvents = True
    End If
End Sub
```

### Exercise 5: Create Summary Dashboard on Summary Sheet

```vba
Sub CreateSummary()
    Dim wsSum As Worksheet
    Set wsSum = Worksheets("Summary")
    wsSum.Cells.Clear
    
    With wsSum
        .Range("A1").Value = "Printing Company Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Total Orders"
        .Range("B3").Formula = "=COUNTA(Orders!A:A)-1"
        
        .Range("A4").Value = "Total Revenue"
        .Range("B4").Formula = "=SUM(Orders!F:F)"
        
        .Range("A5").Value = "Average Order Value"
        .Range("B5").Formula = "=AVERAGE(Orders!F:F)"
        
        .Range("A7").Value = "Orders by Product"
        .Range("A8").Value = "Product"
        .Range("B8").Value = "Count"
        .Range("C8").Value = "Revenue"
        
        ' Pivot-like summary using dictionary (late binding)
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim rng As Range
        Set rng = Worksheets("Orders").Range("C2:F" & Worksheets("Orders").Cells(Rows.Count, "C").End(xlUp).Row)
        
        Dim row As Range
        For Each row In rng.Rows
            Dim prod As String
            prod = row.Cells(1, 1).Value
            If dict.exists(prod) Then
                dict(prod) = dict(prod) + row.Cells(1, 4).Value  ' Add total price
            Else
                dict(prod) = row.Cells(1, 4).Value
            End If
        Next row
        
        Dim key As Variant
        Dim i As Integer: i = 9
        For Each key In dict.keys
            .Cells(i, 1).Value = key
            .Cells(i, 3).Value = dict(key)
            i = i + 1
        Next key
    End With
    
    MsgBox "Summary dashboard created!"
End Sub
```

### Exercise 6: Simple UserForm for Adding Orders

1. In VBA Editor: Insert > UserForm.
2. Add Labels, TextBoxes (txtCustomer, txtProduct, txtQuantity), and a CommandButton (cmdAdd).
3. Double-click cmdAdd and paste:

```vba
Private Sub cmdAdd_Click()
    Dim ws As Worksheet
    Set ws = Worksheets("Orders")
    
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    With ws
        .Cells(lastRow, 1).Value = lastRow - 1
        .Cells(lastRow, 2).Value = txtCustomer.Text
        .Cells(lastRow, 3).Value = txtProduct.Text
        .Cells(lastRow, 4).Value = Val(txtQuantity.Text)
        .Cells(lastRow, 5).Value = GetPrice(txtProduct.Text)
        .Cells(lastRow, 6).Value = Val(txtQuantity.Text) * GetPrice(txtProduct.Text)
        .Cells(lastRow, 7).Value = Date
        .Cells(lastRow, 8).Value = "Pending"
    End With
    
    MsgBox "Order added!"
    Unload Me
End Sub
```

4. Create a macro to show the form:

```vba
Sub ShowAddOrderForm()
    UserForm1.Show
End Sub
```
