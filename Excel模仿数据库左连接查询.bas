Attribute VB_Name = "Excelģ�����ݿ������Ӳ�ѯ"

Sub ShowLeftJoinForm()

    LeftJoinForm.Show vbModeless
  
End Sub


Private Sub TestLeftJoinFun()
    
    Dim startCellName As String, keyRangeAName As String, keyRangeBName As String, valueRangeName As String
    Dim sheetCName As String, sheetAName As String, sheetBName As String
    
'=====================================================================
    ' ���д���C���ӵ�Ԫ��A2��ʼд
    ' Key �ڱ�A�� A2~A161 ����
    ' Key �ڱ�B�� A2~A264 ����
    ' Value �ڱ�B�� B2~B264 ����

    sheetCName = "��C"
    sheetAName = "��A"
    sheetBName = "��B"
    
    startCellName = "A2"
    keyRangeAName = "A2:A161"
    keyRangeBName = "A2:A264"
    valueRangeName = "B2:B264"
'=====================================================================
 
    Dim startCell As Range, keyRangeA As Range, keyRangeB As Range, valueRange As Range
    
    Set startCell = Worksheets(sheetCName).Range(startCellName)
    Set keyRangeA = Worksheets(sheetAName).Range(keyRangeAName)
    Set keyRangeB = Worksheets(sheetBName).Range(keyRangeBName)
    Set valueRange = Worksheets(sheetBName).Range(valueRangeName)
    
    Call LeftJoinFun(startCell, keyRangeA, keyRangeB, valueRange)
    
    Set startCell = Nothing
    Set keyRangeA = Nothing
    Set keyRangeB = Nothing
    Set valueRange = Nothing
    
End Sub

'
' ���ݱ�A�е�Keyֵ���ڱ�B�в��Ҷ�Ӧ��Value,���Ѽ�ֵ��Ӧ���д���C
'
' @param { Range } startCell �ӱ�C��ָ����Ԫ��ʼд��
' @param { Range } keyRangeA Key�ڱ�A�е�����(������)
' @param { Range } keyRangeB Key�ڱ�B�е�����
' @param { Range } valueRange Value�ڱ�B�е�����Ŀ������
'
' @return ��
'

Public Function LeftJoinFun(ByRef startCell As Range, ByVal keyRangeA As Range, ByVal keyRangeB As Range, ByVal valueRange As Range)

    Application.ScreenUpdating = False

    ' ���һ�� Key - Value,�������Ԫ���ǰ��� Key �� Value ��ֵ�Ե�����
    Dim list As New Collection
    ' ��B�� Value �к� Key �еļ��
    Dim dist As Integer
    dist = valueRange.Column - keyRangeB.Column
    
    For Each KA In keyRangeA
    
        ' ��������������B���ҵ��Ķ�Ӧֵ
        Dim i As Integer
        i = 0
        ' ���һ�� Key - Value
        Dim mapArray(0 To 1) As Variant
        
        For Each KB In keyRangeB
        
            If KB.Value Like KA.Value Then
                mapArray(0) = KB.Value
                mapArray(1) = KB.Offset(0, dist)
                list.Add item:=mapArray
                i = i + 1
            End If

        Next
        
        ' ����ڱ�B���Ҳ�������ȻҪ�ѿ�ֵд�뼯�ϣ��������ݿ�������Ӳ�ѯ����ֵ�� # ��ʾ
        If i = 0 Then
            mapArray(0) = KA.Value
            mapArray(1) = "#"
            list.Add item:=mapArray
        End If
        
    Next

      ' ���ü�ֵ���ڱ�C��д������
    Dim resultRangeC As Range
    Set resultRangeC = startCell.Resize(list.Count, 2)
    
    ' ��������x����д������ĵ�Ԫ��y����list�����е�Ԫ��
    Dim x As Integer, y As Integer
    x = 1
    y = 1

    For Each ran In resultRangeC

        ' ����ĵ�һ�У�д��Key
        If x Mod 2 = 1 Then
            ran.Value = list.item(y)(0)
        Else
        ' ����ĵڶ��У�д��Value
            ran.Value = list.item(y)(1)
            y = y + 1
        End If
        x = x + 1
        
    Next

    Application.ScreenUpdating = True
    
    Set resultRangeC = Nothing
    
End Function










