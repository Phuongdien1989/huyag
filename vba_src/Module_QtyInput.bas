Option Explicit

Private Const DATA_SHEET_NAME As String = "DATA"

Public Sub ShowQtyInputForm()
    On Error GoTo ErrHandler

    frmQtyInput.Show

    Exit Sub
ErrHandler:
    MsgBox "Không thể mở form nhập sản lượng: " & Err.Description, vbExclamation
End Sub

Public Function GetDataSheet() As Worksheet
    On Error GoTo ErrHandler

    Set GetDataSheet = ThisWorkbook.Worksheets(DATA_SHEET_NAME)

    Exit Function
ErrHandler:
    Err.Raise vbObjectError + 1000, "GetDataSheet", _
              "Không tìm thấy sheet '" & DATA_SHEET_NAME & "'."
End Function

Public Function NextDataRow(ByVal targetSheet As Worksheet, Optional ByVal columnIndex As Long = 1) As Long
    Dim lastRow As Long

    If Application.WorksheetFunction.CountA(targetSheet.Columns(columnIndex)) = 0 Then
        NextDataRow = 1
        Exit Function
    End If

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, columnIndex).End(xlUp).Row

    If lastRow = targetSheet.Rows.Count And Len(targetSheet.Cells(lastRow, columnIndex).Value) = 0 Then
        lastRow = 0
    End If

    NextDataRow = lastRow + 1
End Function

Public Function GetItemCodeList() As Variant
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim itemCode As String

    Set ws = GetDataSheet()
    lastRow = NextDataRow(ws, 1) - 1

    Set dict = CreateObject("Scripting.Dictionary")

    For rowIndex = 1 To lastRow
        itemCode = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        If Len(itemCode) > 0 Then
            If Not dict.Exists(itemCode) Then
                dict.Add itemCode, itemCode
            End If
        End If
    Next rowIndex

    If dict.Count = 0 Then
        GetItemCodeList = Array()
        Exit Function
    End If

    GetItemCodeList = dict.Keys

    Exit Function
ErrHandler:
    MsgBox "Không thể tải danh sách Mã hàng: " & Err.Description, vbExclamation
    GetItemCodeList = Array()
End Function

Public Sub AppendItemQuantity(ByVal itemCode As String, ByVal quantity As Double)
    Dim screenUpdatingState As Boolean
    Dim enableEventsState As Boolean
    Dim calculationState As XlCalculation

    screenUpdatingState = Application.ScreenUpdating
    enableEventsState = Application.EnableEvents
    calculationState = Application.Calculation

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim targetRow As Long

    Set ws = GetDataSheet()
    targetRow = NextDataRow(ws, 1)

    ws.Cells(targetRow, 1).Value = itemCode
    ws.Cells(targetRow, 2).Value = quantity

CleanExit:
    Application.ScreenUpdating = screenUpdatingState
    Application.EnableEvents = enableEventsState
    Application.Calculation = calculationState

    Exit Sub
ErrHandler:
    MsgBox "Không thể ghi sản lượng: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
