VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LeftJoinForm 
   Caption         =   "左连接查询向导"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   OleObjectBlob   =   "LeftJoinForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "LeftJoinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ConfirmBtn_Click()
    Dim startCell As Range, keyRangeA As Range, keyRangeB As Range, valueRange As Range
    
    On Error GoTo a
    
    Set startCell = Worksheets(sheetCList.Value).Range(startCellText.Value)
    Set keyRangeA = Worksheets(sheetAList.Value).Range(keyRangeAText.Value)
    Set keyRangeB = Worksheets(sheetBList.Value).Range(keyRangeBText.Value)
    Set valueRange = Worksheets(sheetBList.Value).Range(valueRangeText.Value)

    Call LeftJoinFun(startCell, keyRangeA, keyRangeB, valueRange)

    Set startCell = Nothing
    Set keyRangeA = Nothing
    Set keyRangeB = Nothing
    Set valueRange = Nothing
    
    Exit Sub
    
a:     MsgBox "存在空值/非法值！请确保所有信息都已正确填写！"

End Sub


Private Sub ExitBtn_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
' 窗体初始化时，设置复选框的 List 属性为工作表名称的集合

    Dim sheetsArray() As String
    ReDim sheetsArray(ActiveWorkbook.Worksheets.Count - 1)
    
    Dim i As Integer
    i = 0
    
    ' 将活动工作簿中，所有工作表的名称放入数组
    For Each Sheet In ActiveWorkbook.Worksheets
        sheetsArray(i) = Sheet.Name
        i = i + 1
    Next
    
    sheetCList.list = sheetsArray
    sheetAList.list = sheetsArray
    sheetBList.list = sheetsArray
    
End Sub

