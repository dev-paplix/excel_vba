
# Excel VBA Advanced Training: Printing Company Dataset and Exercises

This Markdown file contains the same sample dummy dataset for a printing company, enhanced for advanced exercises. The focus is on **advanced VBA techniques** including:

- Class Modules for custom objects
- Dictionary and Collection for data management
- ADO/ADO DB for external data (simulated)
- API-like interaction (JSON parsing simulation)
- File/Folder operations (export/import)
- Ribbon customization (basic XML)
- Add-in concepts and error logging

## Sample Dummy Dataset for Printing Company

Setup the workbook as follows (save as `.xlsm`):

1. **Sheet: Orders** – Main data (paste into A1)

| Order ID | Customer Name    | Product Type   | Quantity | Price per Unit | Total Price       | Order Date  | Status    | Notes     |
|----------|------------------|----------------|----------|----------------|-------------------|-------------|-----------|-----------|
| 1        | John Doe        | Brochures     | 500     | 0.25          | =D2*E2           | 2025-01-15 | Pending   |           |
| 2        | Jane Smith      | Flyers        | 1000    | 0.15          | =D3*E3           | 2025-02-20 | Completed | Urgent    |
| 3        | Acme Corp       | Business Cards| 2000    | 0.10          | =D4*E4           | 2025-03-05 | Pending   |           |
| 4        | Bob Johnson     | Posters       | 100     | 2.50          | =D5*E5           | 2025-04-10 | Completed |           |
| 5        | Alice Wonderland| Banners       | 50      | 5.00          | =D6*E6           | 2025-05-15 | Pending   | Large format |
| 6        | Tech Inc        | Brochures     | 750     | 0.20          | =D7*E7           | 2025-06-20 | Completed |           |
| 7        | Foodies Ltd     | Menus         | 300     | 0.40          | =D8*E8           | 2025-07-25 | Pending   |           |
| 8        | Sports Gear     | Flyers        | 1500    | 0.12          | =D9*E9           | 2025-08-30 | Completed |           |
| 9        | Book Club       | Booklets      | 200     | 1.00          | =D10*E10         | 2025-09-05 | Pending   |           |
| 10       | Event Planners  | Invitations   | 400     | 0.30          | =D11*E11         | 2025-10-10 | Completed |           |

2. **Sheet: Products** – Price list (A1:B7)

| Product Type    | Standard Price | Discount Threshold | Discount % |
|-----------------|----------------|--------------------|------------|
| Brochures      | 0.25          | 1000               | 10        |
| Flyers         | 0.15          | 2000               | 15        |
| Business Cards | 0.10          | 3000               | 20        |
| Posters        | 2.50          | 200                | 5         |
| Banners        | 5.00          | 100                | 8         |
| Others         | 0.50          | 500                | 10        |

3. **Sheet: Summary** – Will be populated by code.

## Advanced VBA Exercises

### Exercise 1: Create a Class Module - COrder

1. Insert > Class Module → Name it **COrder**

```vba
' COrder Class
Private pOrderID As Long
Private pCustomer As String
Private pProduct As String
Private pQuantity As Long
Private pPrice As Double
Private pTotal As Double
Private pDate As Date
Private pStatus As String

Public Property Get OrderID() As Long: OrderID = pOrderID: End Property
Public Property Let OrderID(value As Long): pOrderID = value: End Property

Public Property Get Customer() As String: Customer = pCustomer: End Property
Public Property Let Customer(value As String): pCustomer = value: End Property

Public Property Get Product() As String: Product = pProduct: End Property
Public Property Let Product(value As String): pProduct = value: End Property

Public Property Get Quantity() As Long: Quantity = pQuantity: End Property
Public Property Let Quantity(value As Long): pQuantity = value: End Property

Public Property Get Price() As Double: Price = pPrice: End Property
Public Property Let Price(value As Double): pPrice = value: End Property

Public Property Get Total() As Double
    Total = pQuantity * pPrice
End Property

Public Property Get OrderDate() As Date: OrderDate = pDate: End Property
Public Property Let OrderDate(value As Date): pDate = value: End Property

Public Property Get Status() As String: Status = pStatus: End Property
Public Property Let Status(value As String): pStatus = value: End Property

Public Sub CalculateDiscountedPrice(wsProducts As Worksheet)
    Dim found As Range
    Set found = wsProducts.Columns("A").Find(What:=pProduct, LookAt:=xlWhole)
    If Not found Is Nothing Then
        If pQuantity >= found.Offset(0, 2).Value Then
            pPrice = found.Offset(0, 1).Value * (1 - found.Offset(0, 3).Value / 100)
        Else
            pPrice = found.Offset(0, 1).Value
        End If
    End If
End Sub
```

### Exercise 2: Load All Orders into Collection of COrder Objects

```vba
Sub LoadOrdersToCollection()
    Dim colOrders As New Collection
    Dim ws As Worksheet: Set ws = Worksheets("Orders")
    Dim wsProd As Worksheet: Set wsProd = Worksheets("Products")
    
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim ord As New COrder
        With ord
            .OrderID = ws.Cells(i, 1).Value
            .Customer = ws.Cells(i, 2).Value
            .Product = ws.Cells(i, 3).Value
            .Quantity = ws.Cells(i, 4).Value
            .CalculateDiscountedPrice wsProd
            .OrderDate = ws.Cells(i, 7).Value
            .Status = ws.Cells(i, 8).Value
        End With
        colOrders.Add ord
    Next i
    
    ' Example: Calculate total revenue with discounts
    Dim totalRev As Double
    Dim o As COrder
    For Each o In colOrders
        totalRev = totalRev + o.Total
    Next o
    
    MsgBox "Total Revenue with Discounts: " & Format(totalRev, "Currency")
End Sub
```

### Exercise 3: Advanced Dictionary Grouping + Export to JSON-like Text File

```vba
Sub ExportSummaryToFile()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet: Set ws = Worksheets("Orders")
    Dim lastRow As Long: lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim prod As String: prod = ws.Cells(i, 3).Value
        Dim rev As Double: rev = ws.Cells(i, 6).Value
        
        If dict.Exists(prod) Then
            dict(prod) = dict(prod) + rev
        Else
            dict(prod) = rev
        End If
    Next i
    
    ' Create JSON-like string
    Dim json As String
    json = "{" & vbCrLf & "  ""summary"": {" & vbCrLf
    
    Dim key As Variant
    For Each key In dict.Keys
        json = json & "    """ & key & """: " & dict(key) & "," & vbCrLf
    Next key
    
    json = Left(json, Len(json) - 3) & vbCrLf & "  }" & vbCrLf & "}"
    
    ' Export to file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Printing_Summary.json"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, json
    Close #fileNum
    
    MsgBox "Summary exported to: " & filePath
End Sub
```

### Exercise 4: Advanced Error Logging Class

Create class **CErrorLogger**

```vba
' CErrorLogger Class
Private logFile As String

Public Sub Init(logPath As String)
    logFile = logPath
End Sub

Public Sub LogError(procName As String, errNum As Long, errDesc As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, Now & " | " & procName & " | Error " & errNum & ": " & errDesc
    Close #fileNum
End Sub
```

Usage example in a macro:

```vba
Sub SafeMacro()
    Dim logger As New CErrorLogger
    logger.Init ThisWorkbook.Path & "\error_log.txt"
    
    On Error GoTo ErrHandler
    ' Your risky code here
    Range("NonExistentRange").Value = 1
    Exit Sub
    
ErrHandler:
    logger.LogError "SafeMacro", Err.Number, Err.Description
    MsgBox "An error occurred. Logged."
End Sub
```

### Exercise 5: Dynamic Ribbon Button (Custom UI)

1. Export Ribbon XML (use Custom UI Editor tool or manually):
   - Add to workbook via Custom UI Editor.

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="PrintingGroup" label="Printing Tools">
          <button id="btnSummary" label="Create Summary" imageMso="HappyFace" onAction="CreateAdvancedSummary"/>
          <button id="btnExport" label="Export JSON" imageMso="FileSaveAs" onAction="ExportSummaryToFile"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

Callback:

```vba
Sub CreateAdvancedSummary(control As IRibbonControl)
    Call CreateDynamicSummary
End Sub
```

### Exercise 6: Dynamic Pivot-like Summary with Sorting

```vba
Sub CreateDynamicSummary()
    Dim wsSum As Worksheet
    Set wsSum = Worksheets("Summary")
    wsSum.Cells.Clear
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    ' ... (same grouping as Exercise 3)
    
    ' Sort dictionary by value descending
    Dim sortedKeys() As Variant
    ReDim sortedKeys(0 To dict.Count - 1)
    
    Dim keys As Variant: keys = dict.Keys
    Dim values As Variant: values = dict.Items
    
    ' Simple bubble sort
    Dim i As Long, j As Long, tempKey As Variant, tempVal As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If values(i) < values(j) Then
                tempVal = values(i): values(i) = values(j): values(j) = tempVal
                tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            End If
        Next j
    Next i
    
    ' Output sorted
    wsSum.Range("A1").Value = "Product"
    wsSum.Range("B1").Value = "Revenue"
    For i = 0 To UBound(keys)
        wsSum.Cells(i + 2, 1).Value = keys(i)
        wsSum.Cells(i + 2, 2).Value = values(i)
    Next i
    
    wsSum.Columns("A:B").AutoFit
End Sub
```

### Final Advanced Tips
- Use **Class Modules** to model real-world entities.
- Prefer **Collections/Dictionaries** over arrays for dynamic data.
- Always implement **robust error handling and logging**.
- Export data for integration (JSON, CSV).
- Customize the **Ribbon** for professional tools.
- Consider packaging as an **Add-in** (.xlam) for reuse.

These exercises will elevate your VBA skills to advanced/professional level. Practice building reusable, maintainable code!
```

