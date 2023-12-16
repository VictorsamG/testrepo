VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntry1 
   Caption         =   "Sales Entry"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   OleObjectBlob   =   "DataEntry1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataEntry1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public week
Public lastRow As Long

Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub btnEdit_Click()
Dim iRow As Long
Dim jRow As Long
Dim sheetRow As Long
Dim week As Range
    If Me.ListBox1.ListIndex < 0 Then Exit Sub
    Call Clear_Inputs
    iRow = Me.ListBox1.ListIndex
    jRow = Me.ListBox1.List(iRow, 0)
    sheetRow = jRow + 1
    Me.txtsRow = sheetRow
    MsgBox sheetRow
    Me.cmbMonths.Value = Format(Me.ListBox1.List(iRow, 1), "mmmm")
        Me.cmbWeeks.Clear
        Me.txtEoM = ""
        Sheet2.Range("O2").Value = Me.cmbMonths.Value
            For Each week In ThisWorkbook.Names("MonthWeeks").RefersToRange
                Me.cmbWeeks.AddItem week
            Next week
        Me.txtEoM = Sheet2.Range("R2").Value
    Me.cmbWeeks.Value = Me.ListBox1.List(iRow, 2)
    Me.cmbSalesRep.Value = Me.ListBox1.List(iRow, 3)
    Me.cmbChannel.Value = Me.ListBox1.List(iRow, 4)
    Me.txtVenues.Value = Me.ListBox1.List(iRow, 5)
    Me.cmbProducts.Value = Me.ListBox1.List(iRow, 6)
    Me.txtQty.Value = Me.ListBox1.List(iRow, 7)
    Me.txtAmount.Value = Me.ListBox1.List(iRow, 8)
End Sub

Private Sub btnSave_Click()
    Call Save_Records
    ListBox1.RowSource = "InputSheet!A2:I2"
    Me.txtTTqty = InputSheet.Range("L1").Value
    Me.txtTTam = InputSheet.Range("O1").Value
End Sub

Private Sub btnInsert_Click()
    Call Insert_Entry
End Sub

Private Sub cmbMonths_Change()
Me.cmbWeeks = ""
End Sub

Private Sub cmbMonths_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Me.cmbWeeks.Clear
Me.txtEoM = ""

Sheet2.Range("O2").Value = Me.cmbMonths.Value
    For Each week In ThisWorkbook.Names("MonthWeeks").RefersToRange
        Me.cmbWeeks.AddItem week
    Next week
Me.txtEoM = Sheet2.Range("R2").Value
Cancel = False
End Sub

Private Sub cmbProducts_Change()
If Me.cmbProducts.Value <> "" Then
    Me.txtPrice = Application.WorksheetFunction.VLookup(Me.cmbProducts.Value, _
                Sheet5.Range("F:G"), 2, False)
Else
    Me.txtPrice = ""
End If
End Sub



Private Sub btnDelete_Click()
Dim i As Integer
Dim cell As Range
Dim lastRow As Long
Dim msgrslt
    If ListBox1.ListIndex < 0 Then Exit Sub
    msgrslt = MsgBox("Are you sure you want to delete this entry?", vbYesNo, "Confirm Delete Entry")
    If msgrslt = vbNo Then Exit Sub
    If msgrslt = vbYes Then
        InputSheet.Range(ListBox1.ListIndex + 2 & ":" & ListBox1.ListIndex + 2).EntireRow.Delete
        lastRow = InputSheet.Range("B100000").End(xlUp).Row
        If lastRow > 1 Then
            i = 1
            For Each cell In InputSheet.Range("A2:A" & lastRow)
                cell = i
                i = i + 1
            Next cell
        Else
            InputSheet.Range("A2").Value = 1
        End If
        If Me.txtsRow <> "" Then Call Clear_Inputs
        MsgBox "The entry has been deleted"
    End If
End Sub

Private Sub ListBox1_Click()
    Me.btnDelete.Visible = True
    Me.btnEdit.Visible = True
End Sub

Private Sub txtAmount_Change()
If Not IsNumeric(Me.txtAmount.Value) And Me.txtAmount.Value <> "" Then
    MsgBox "Please Enter A Numeric Value"
    Me.txtQty = ""
End If
End Sub

Private Sub txtQty_Change()
If Not IsNumeric(Me.txtQty.Value) And Me.txtQty.Value <> "" Then
    MsgBox "Please Enter A Numeric Value"
    Me.txtQty = ""
Else
    If Me.txtQty = "" Then
        Me.txtAmount = ""
        Exit Sub
    Else
        If Me.txtPrice = "" Then
            Me.txtAmount = ""
            Exit Sub
        End If
    End If
     Me.txtAmount = Format(Me.txtQty * Me.txtPrice, "#,0.00")
End If

End Sub

Private Sub UserForm_Activate()
    
    lastRow = InputSheet.Range("B100000").End(xlUp).Row
    
    With Me.ListBox1
        'If Me.ListBox1.Selected = False Then
            Me.btnDelete.Visible = False
            Me.btnEdit.Visible = False
        'End If
        .ColumnCount = 9
        .ColumnWidths = "20;70;40;60;80;40;80;40;60"
        .ColumnHeads = True
        If lastRow > 1 Then
            .RowSource = "InputSheet!A2:I" & lastRow
        Else
            .RowSource = "InputSheet!A2:I2"
        End If
    End With
    Me.txtTTqty = InputSheet.Range("L1").Value
    Me.txtTTam = InputSheet.Range("O1").Value
End Sub

Private Sub Insert_Entry()
Dim i As Integer
Dim cell As Range
    If Me.cmbMonths = "" Then
    MsgBox "Please Select a Month"
    Exit Sub
    End If
    If Me.cmbWeeks = "" Then
    MsgBox "Please Select a Week"
    Exit Sub
    End If
    If Me.cmbSalesRep = "" Then
    MsgBox "Please Select a SalesRep"
    Exit Sub
    End If
    If Me.cmbChannel = "" Then
    MsgBox "Please Select a Sales Channel"
    Exit Sub
    End If
    If Me.cmbProducts = "" Then
    MsgBox "Please Select a Product"
    Exit Sub
    End If
    If Me.txtQty = "" Then
    MsgBox "Please Enter Quantity"
    Exit Sub
    End If
    If Me.txtAmount = "" Then
    MsgBox "Please Enter Sales Amount"
    Exit Sub
    End If
    With InputSheet
        If Me.txtsRow.Value = "" Then
            lastRow = .Range("B100000").End(xlUp).Row + 1
        Else
            lastRow = Me.txtsRow.Value
        End If
        .Range("B" & lastRow).Value = Me.txtEoM.Value
        .Range("C" & lastRow).Value = Me.cmbWeeks.Value
        .Range("D" & lastRow).Value = Me.cmbSalesRep.Value
        .Range("E" & lastRow).Value = Me.cmbChannel.Value
        .Range("F" & lastRow).Value = Me.txtVenues.Value
        .Range("G" & lastRow).Value = Me.cmbProducts.Value
        .Range("H" & lastRow).Value = Me.txtQty.Value
        .Range("I" & lastRow).Value = Me.txtAmount.Value
        i = 1
        lastRow = .Range("B100000").End(xlUp).Row
        For Each cell In .Range("A2:A" & lastRow)
            cell = i
            i = i + 1
        Next cell
    End With
    
    If lastRow > 1 Then
        Me.ListBox1.RowSource = "InputSheet!A2:I" & lastRow
    Else
        Me.ListBox1.RowSource = "InputSheet!A2:I2"
    End If
    Me.txtTTqty = InputSheet.Range("L1").Value
    Me.txtTTam = InputSheet.Range("O1").Value
    Me.cmbProducts.Value = ""
    Me.txtQty = ""
    Me.txtAmount = ""
    Me.cmbChannel.Value = ""
    Me.txtVenues = ""
    Me.txtsRow = ""
End Sub
Private Sub Clear_Inputs()
    Me.cmbWeeks.Value = ""
    Me.cmbMonths.Value = ""
    Me.cmbSalesRep.Value = ""
    Me.cmbChannel.Value = ""
    Me.txtPrice = ""
    Me.cmbProducts.Value = ""
    Me.txtEoM = ""
    Me.txtQty = ""
    Me.txtAmount = ""
    Me.txtVenues = ""
    Me.txtsRow = ""
End Sub
