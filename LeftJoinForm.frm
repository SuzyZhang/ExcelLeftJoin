VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LeftJoinForm 
   Caption         =   "�����Ӳ�ѯ��"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   OleObjectBlob   =   "LeftJoinForm.frx":0000
   StartUpPosition =   1  '����������
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
    
a:     MsgBox "���ڿ�ֵ/�Ƿ�ֵ����ȷ��������Ϣ������ȷ��д��"

End Sub


Private Sub ExitBtn_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
' �����ʼ��ʱ�����ø�ѡ��� List ����Ϊ���������Ƶļ���

    Dim sheetsArray() As String
    ReDim sheetsArray(ActiveWorkbook.Worksheets.Count - 1)
    
    Dim i As Integer
    i = 0
    
    ' ����������У����й���������Ʒ�������
    For Each Sheet In ActiveWorkbook.Worksheets
        sheetsArray(i) = Sheet.Name
        i = i + 1
    Next
    
    sheetCList.list = sheetsArray
    sheetAList.list = sheetsArray
    sheetBList.list = sheetsArray
    
End Sub

