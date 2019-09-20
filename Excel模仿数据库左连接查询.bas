Attribute VB_Name = "Excel模仿数据库左连接查询"

Sub ShowLeftJoinForm()

    LeftJoinForm.Show vbModeless
  
End Sub


Private Sub TestLeftJoinFun()
    
    Dim startCellName As String, keyRangeAName As String, keyRangeBName As String, valueRangeName As String
    Dim sheetCName As String, sheetAName As String, sheetBName As String
    
'=====================================================================
    ' 结果写入表C，从单元格A2开始写
    ' Key 在表A的 A2~A161 区域
    ' Key 在表B的 A2~A264 区域
    ' Value 在表B的 B2~B264 区域

    sheetCName = "表C"
    sheetAName = "表A"
    sheetBName = "表B"
    
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
' 根据表A中的Key值，在表B中查找对应的Value,并把键值对应结果写入表C
'
' @param { Range } startCell 从表C的指定单元格开始写入
' @param { Range } keyRangeA Key在表A中的区域(索引区)
' @param { Range } keyRangeB Key在表B中的区域
' @param { Range } valueRange Value在表B中的区域（目标区）
'
' @return 无
'

Public Function LeftJoinFun(ByRef startCell As Range, ByVal keyRangeA As Range, ByVal keyRangeB As Range, ByVal valueRange As Range)

    Application.ScreenUpdating = False

    ' 存放一组 Key - Value,集合里的元素是包含 Key 和 Value 键值对的数组
    Dim list As New Collection
    ' 表B中 Value 列和 Key 列的间隔
    Dim dist As Integer
    dist = valueRange.Column - keyRangeB.Column
    
    For Each KA In keyRangeA
    
        ' 计数器，计数表B中找到的对应值
        Dim i As Integer
        i = 0
        ' 存放一对 Key - Value
        Dim mapArray(0 To 1) As Variant
        
        For Each KB In keyRangeB
        
            If KB.Value Like KA.Value Then
                mapArray(0) = KB.Value
                mapArray(1) = KB.Offset(0, dist)
                list.Add item:=mapArray
                i = i + 1
            End If

        Next
        
        ' 如果在表B中找不到，依然要把空值写入集合，类似数据库的左连接查询。空值用 # 表示
        If i = 0 Then
            mapArray(0) = KA.Value
            mapArray(1) = "#"
            list.Add item:=mapArray
        End If
        
    Next

      ' 设置键值对在表C的写入区域
    Dim resultRangeC As Range
    Set resultRangeC = startCell.Resize(list.Count, 2)
    
    ' 计数器，x计数写入区域的单元格，y计数list集合中的元素
    Dim x As Integer, y As Integer
    x = 1
    y = 1

    For Each ran In resultRangeC

        ' 区域的第一列，写入Key
        If x Mod 2 = 1 Then
            ran.Value = list.item(y)(0)
        Else
        ' 区域的第二列，写入Value
            ran.Value = list.item(y)(1)
            y = y + 1
        End If
        x = x + 1
        
    Next

    Application.ScreenUpdating = True
    
    Set resultRangeC = Nothing
    
End Function










