VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQtyInput 
   Caption         =   "Nhập sản lượng"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.CommandButton btnSave 
      Caption         =   "Ghi"
      Default         =   -1  'True
      Height          =   360
      Left            =   2700
      TabIndex        =   4
      Top             =   2580
      Width           =   1005
   End
   Begin MSForms.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Đóng"
      Height          =   360
      Left            =   3840
      TabIndex        =   5
      Top             =   2580
      Width           =   1005
   End
   Begin MSForms.TextBox txtQuantity 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   4725
   End
   Begin MSForms.Label lblQuantity 
      Caption         =   "Số lượng"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   1005
   End
   Begin MSForms.ComboBox cboItem 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4725
   End
   Begin MSForms.Label lblItem 
      Caption         =   "Mã hàng"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1005
   End
End
Attribute VB_Name = "frmQtyInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrHandler

    Dim itemCode As String
    Dim quantityValue As Double

    If Not InputsAreValid(itemCode, quantityValue) Then
        Exit Sub
    End If

    AppendItemQuantity itemCode, quantityValue
    MsgBox "Đã ghi nhận sản lượng.", vbInformation

    ResetForm

    Exit Sub
ErrHandler:
    MsgBox "Có lỗi khi lưu dữ liệu: " & Err.Description, vbCritical
End Sub

Private Function InputsAreValid(ByRef itemCode As String, ByRef quantityValue As Double) As Boolean
    itemCode = Trim$(CStr(Me.cboItem.Value))

    If Len(itemCode) = 0 Then
        MsgBox "Vui lòng chọn Mã hàng.", vbExclamation
        Me.cboItem.SetFocus
        Exit Function
    End If

    If Not TryParseQuantity(Me.txtQuantity.Value, quantityValue) Then
        MsgBox "Số lượng không hợp lệ. Vui lòng nhập số.", vbExclamation
        Me.txtQuantity.SetFocus
        Me.txtQuantity.SelStart = 0
        Me.txtQuantity.SelLength = Len(Me.txtQuantity.Text)
        Exit Function
    End If

    If quantityValue <= 0 Then
        MsgBox "Số lượng phải lớn hơn 0.", vbExclamation
        Me.txtQuantity.SetFocus
        Exit Function
    End If

    InputsAreValid = True
End Function

Private Sub ResetForm()
    Me.cboItem.Value = vbNullString
    Me.txtQuantity.Value = vbNullString
    Me.cboItem.SetFocus
End Sub

Private Sub UserForm_Initialize()
    LoadItemCodes
End Sub

Private Sub LoadItemCodes()
    On Error GoTo ErrHandler

    Dim itemCodes As Variant

    itemCodes = GetItemCodeList()

    Me.cboItem.Clear

    If HasItems(itemCodes) Then
        Me.cboItem.List = itemCodes
    End If

    Exit Sub
ErrHandler:
    MsgBox "Không thể tải danh sách Mã hàng: " & Err.Description, vbExclamation
End Sub

Private Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo ExitPoint

    If IsArray(values) Then
        HasItems = (LBound(values) <= UBound(values))
    End If

ExitPoint:
End Function

Private Function TryParseQuantity(ByVal rawValue As Variant, ByRef parsedValue As Double) As Boolean
    On Error GoTo ErrHandler

    If Len(Trim$(CStr(rawValue))) = 0 Then
        Exit Function
    End If

    parsedValue = CDbl(rawValue)
    TryParseQuantity = True

    Exit Function
ErrHandler:
    TryParseQuantity = False
End Function
