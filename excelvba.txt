Sub 从一个指定的文本文件中读取每一行数据_如果某行以孙兴华开头_写入Excel工作表中的单元格()
 Dim i, j
 Open "D:\BaiduSyncdisk\跟着孙兴华学习Excel VBA\课件\14.使用VBA批量操作TXT和Excel文件\孙兴华.txt" For Input As #1
 i = 1
 '代号1这个文件没有到达末尾时
 Do While Not EOF(1)
 '从1号文件中读取一行文本，把这行文本做为一个字符串保存到s变量中
 Line Input #1, j
 '进行判断，以孙兴华开头的写入单元格
 If Left(j, 3) = "孙兴华" Then
 Range("A" & i) = j
 i = i + 1
 End If
 Loop
 Close #1
End Sub


Sub 用户自己选择文件()
    Dim filePath As String
    Dim i As Long
    Dim lineText As String
'    弹出文件选择对话框，用户选择要打开的文本文件。过滤器限定为 .txt 文件
    ' 用户选择文件路径
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt")
'如果用户取消了文件选择（返回“False”），则退出宏
    If filePath = "False" Then Exit Sub
'    启用错误处理机制，若发生错误会跳转到 ErrorHandler 标签处。
    On Error GoTo ErrorHandler
'    以只读模式打开用户选择的文本文件，并将文件句柄分配给 #1。
    Open filePath For Input As #1
    i = 1
    ' 读取文件内容
    '    使用 EOF(1) 来判断文件是否读取到末尾。如果未到达末尾，则继续读取
    Do While Not EOF(1)
'        读取文件中的一行文本并存储到 lineText 变量中
        Line Input #1, lineText

        ' 筛选以 "孙兴华" 开头的行
        If Left(lineText, 3) = "孙兴华" Then
'        将符合条件的行写入 Excel 中，从第 i 行开始
            Range("A" & i) = lineText
            i = i + 1
        End If
    Loop

    Close #1
'    弹出提示框，告知用户文件处理已完成
    MsgBox "文件内容已处理完成！", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "文件操作出错，请检查路径或内容格式。", vbExclamation
    Close #1
End Sub


'这段代码会遍历工作簿中的所有工作表，对于每张工作表，从第 2 行开始读取 A 列和 B 列的数据
'并将每一行的数据（A 列和 B 列的内容以逗号分隔）写入指定的文本文件 李小龙.txt

Sub 多张工作表同时写入一个文件()
Dim i, s1
Open "D:\BaiduSyncdisk\跟着孙兴华学习Excel VBA\课件\14.使用VBA批量操作TXT和Excel文件\李小龙.txt" For Output As #1
For Each s1 In Sheets
 '每张工作表从第2行开始扫描每一行
 i = 2
 Do While s1.Range("A" & i) <> ""
 Print #1, s1.Range("A" & i); ","; s1.Range("B" & i)
 i = i + 1
 Loop
Next
Close #1
End Sub



'这段代码首先从 姓名.txt 和 功夫.txt 两个文本文件中读取内容，并依次将每行内容写入 Excel 的 A 列。
'然后，将 Excel 中 A 列的内容写入一个新的文本文件 合并.txt。
'最后，关闭所有文件。

Sub 多文件的读取与写入()
Dim i
Open "C:\Users\孙艺航\Desktop\多文件打开写入\姓名.txt" For Input As #1
Open "C:\Users\孙艺航\Desktop\多文件打开写入\功夫.txt" For Input As #2
i = 1
Do While Not EOF(1) Or Not EOF(2)
' 如果文件 1 (#1) 尚未读取完毕
 If Not EOF(1) Then
 Line Input #1, s
 Range("A" & i) = s
 i = i + 1
 End If
' 如果文件 2 (#2) 尚未读取完毕
 If Not EOF(2) Then
 Line Input #2, s
 Range("A" & i) = s
 i = i + 1
 End If
Loop
Close #1: Close #2
Open "C:\Users\孙艺航\Desktop\多文件打开写入\合并.txt" For Output As #3
i = 1
Do While Range("A" & i) <> ""
'将 A 列中的当前行内容写入 合并.txt 文件
 Print #3, Range("A" & i)
 i = i + 1
Loop
'关闭新创建的文本文件 合并.txt
Close #3
End Sub


Sub 遍历所有txt文件()
Dim 文件
'运行Dir函数得到第1个文件的名字
文件 = Dir("C:\Users\孙艺航\Desktop\多文件打开写入\")
'如果读到的文件不是空字符串，就证明这是一个有效文件
Do While 文件 <> ""
 '这里可以对文件进行打开和读取操作
 文件 = Dir '再次运行Dir就读到下一个文件名
Loop
End Sub

'目的是批量读取一个文件夹中的多个 .txt 文件，并将每个文件的内容按行读取，分割并写入到 Excel 工作表中
'每个文件的内容会写入一个新的工作表，并且每个工作表的名称会使用文件名（不包含路径部分）

Sub 批量读取一个文件夹中的多个txt文件()
Dim 文件
'运行Dir函数得到第1个文件的名字
文件 = Dir("C:\Users\孙艺航\Desktop\txt\")
'如果读到的文件不是空字符串，就证明这是一个有效文件
Do While 文件 <> ""
 '这里可以对文件进行打开和读取操作
 Call 读取多个txt文件("C:\Users\孙艺航\Desktop\txt\" & 文件)
 文件 = Dir '再次运行Dir就读到下一个文件名
Loop
End Sub
'读取【带路径的文件名】变量中存储的文件
'取出每行国家名称和确诊人数，写入工作表
Sub 读取多个txt文件(带路径的文件名)
Dim i, w1, x
Set w1 = Worksheets.Add
'设置新工作表的名称为文件名（不包含路径）
'InStrRev 函数用于查找文件路径中的最后一个反斜杠（\）
'然后使用 Mid 函数提取文件名部分
w1.Name = Mid(带路径的文件名, InStrRev(带路径的文件名, "\") + 1)
Open 带路径的文件名 For Input As #1
i = 1
Do While Not EOF(1)
 Line Input #1, x
w1.Range("A" & i) = Split(x, ",")(0)
 w1.Range("B" & i) = Split(x, ",")(1)
 w1.Range("C" & i) = Split(x, ",")(2)
 i = i + 1
Loop
Close #1
End Sub


遍历指定文件夹下的所有 Excel 文件，并逐一打开、处理后关闭每个文件

Sub 遍历文件夹下Excel文件()
Dim w1
文件 = Dir("C:\Users\孙艺航\Desktop\excel\")
Do While 文件 <> ""
'使用 Workbooks.Open 方法打开当前文件，并将其引用赋值给变量 w1
 Set w1 = Workbooks.Open("C:\Users\孙艺航\Desktop\excel\" & 文件)
 '此处可以处理当前打开的工作簿
 
 w1.Close
 文件 = Dir
Loop
End Sub


遍历指定文件夹下的所有 Excel 文件，并将每个文件中第一个工作表复制到当前运行代码的工作簿中，同时将复制的工作表重命名为源文件的文件名（不包括后缀）

Sub 遍历文件夹下Excel文件()
'暂时关闭屏幕更新
Excel.Application.ScreenUpdating = False
Dim w1
文件 = Dir("C:\Users\孙艺航\Desktop\excel\")
Do While 文件 <> ""
 Set w1 = Workbooks.Open("C:\Users\孙艺航\Desktop\excel\" & 文件)
 '打开文件并复制第1张表，放在我这个写代码的工作簿里，有几张表就在表后面粘贴
 w1.Sheets(1).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
 '刚复制的这张表的表名就是w1那个变量的文件名（不要后缀）
'使用 Split 函数按点分割文件名，并提取分割后的第一个部分（即文件名的主干部分）
 ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = Split(w1.Name, ".")(0)
 w1.Close
 文件 = Dir
Loop
Excel.Application.ScreenUpdating = True
End Sub


遍历指定文件夹中的所有 Excel 文件，并将每个文件的所有工作表复制到当前运行代码的工作簿中。复制的每个工作表都会被重新命名，格式为：文件名.工作表名
Excel 工作表名称最大长度为 31 个字符。如果 文件名.工作表名 超过 31 个字符，代码会报错。

Sub 遍历文件夹下Excel文件()
Excel.Application.ScreenUpdating = False
Dim w1
文件 = Dir("C:\Users\孙艺航\Desktop\多表excel\")
Do While 文件 <> ""
 Set w1 = Workbooks.Open("C:\Users\孙艺航\Desktop\多表excel\" & 文件)
 For Each s1 In w1.Sheets
 '复制s1放到工作表最后面
 s1.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
 '刚复制的这张表的表名就是w1那个变量的文件名（不要后缀）
 ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = Split(w1.Name, ".")(0) & "." & s1.Name
 Next
 w1.Close
 文件 = Dir
Loop
Sheet1.Select
Excel.Application.ScreenUpdating = True
End Sub


遍历指定文件夹下的所有文件，并筛选出以 .xlsx 或 .xls 为扩展名的 Excel 文件，对每个符合条件的文件进行操作

Sub 遍历文件夹下Excel文件()
Dim w1
文件 = Dir("C:\Users\孙艺航\Desktop\excel\")
Do While 文件 <> ""
 'LCase判断是否以xlsx或xls结尾的文件，英文要考虑大小写一致
'LCase(...)：将字符转换为小写，确保文件扩展名判断不受大小写影响
 If LCase(Right(文件, 5)) = ".xlsx" Or LCase(Right(文件, 4)) = ".xls" Then
 Set w1 = Workbooks.Open("C:\Users\孙艺航\Desktop\excel" & 文件)
 '此处可以处理当前打开的工作簿
 
 w1.Close
 End If
 文件 = Dir
Loop
End Sub


目的是逐行扫描某一列，直到找到空白单元格或达到工作表最大行数为止

Sub 遍历行找到工作表最大行数或空白单元格()
    Dim i As Long
    i = 2 ' 从第二行开始
    Do While Range("A" & i) <> "" And i < ActiveSheet.Rows.Count '确保 i 不会超过当前工作表的最大行数（通常为 1,048,576）
        ' 在这里添加需要对每行执行的操作
        Debug.Print "第" & i & "行的值是：" & Range("A" & i).Value
        i = i + 1 ' 增加行号，避免死循环
    Loop
    MsgBox "遍历完成！共扫描到第" & (i - 1) & "行。", vbInformation
End Sub


确定当前工作表中最后一行的行号，并通过消息框显示出来

Sub 确定表格最后一行通过消息框表示()
 Dim r1, i
 '当前工作表所使用的区域
'返回当前活动工作表中实际使用的单元格区域（包括所有非空单元格）
 Set r1 = ActiveSheet.UsedRange
 'Row从第几行开始+总计多少行-1就得到最后一行的位置了
 i = r1.Row + r1.Rows.Count - 1
 MsgBox "最后一行是" & i
End Sub




Sub 在当前工作表中将某列的所有行填充为一个固定值()
Dim i, r1
'返回当前活动工作表的已使用区域（包含所有非空单元格）
Set r1 = ActiveSheet.UsedRange
For i = 1 To r1.Row + r1.Rows.Count - 1
 Range("G" & i) = 520
Next
End Sub


Sub 在当前工作表中将某列的所有行填充为一个固定值()
    Dim i As Long
    Dim r1 As Range

    ' 获取当前使用区域
'如果工作表中只有一个单元格被“使用”，可能是空白的
'检查使用区域的第一个单元格是否为空
    If ActiveSheet.UsedRange.Cells.Count = 1 And IsEmpty(ActiveSheet.UsedRange.Cells(1)) Then
        MsgBox "当前工作表为空！", vbExclamation
        Exit Sub
    End If

    Set r1 = ActiveSheet.UsedRange

    ' 从使用区域的起始行到最后一行
    For i = r1.Row To r1.Row + r1.Rows.Count - 1
        Range("G" & i) = 520
    Next

    MsgBox "填充完成！", vbInformation
End Sub


数组来操作 Excel 工作表的整个 A 列
数据加载：从 Excel 的 A 列读取有数据的范围（如 A1:A100），加载到数组 arr
数组操作：遍历 arr，将数组中的每个值改为 1。
数据回写：将修改后的 arr 写回 Excel 的 A 列对应范围

Sub 配合数组使用优化()
    Dim arr As Variant, i As Long
    Dim lastRow As Long
    ' 找到 A 列最后一个非空单元格的行号
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    ' 加载有数据的部分到数组
    arr = Range("A1:A" & lastRow).Value
    ' 遍历数组并修改数据
    For i = 1 To UBound(arr, 1)
        arr(i, 1) = 1
    Next
    ' 写回修改后的数据
    Range("A1:A" & lastRow).Value = arr
    MsgBox "已经完成"
End Sub


定义源工作表和目标工作表（Sheet1 和 Sheet2）
使用 .CurrentRegion 动态获取以 B2 为起始点的连续区域
清空目标工作表中的所有内容，避免残留数据干扰
将源区域的数据复制到目标工作表中，从 A1 单元格开始粘贴
弹出提示框，告知用户操作完成

Sub 单元格拷贝()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceRange As Range
    Dim targetCell As Range

    ' 定义源工作表和目标工作表
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsTarget = ThisWorkbook.Sheets("Sheet2")

    ' 找到源数据范围
    Set sourceRange = wsSource.Range("B2").CurrentRegion

    ' 定义目标单元格
    Set targetCell = wsTarget.Range("A1")

    ' 清空目标工作表中的区域
    wsTarget.Cells.Clear

    ' 将数据从源复制到目标
    sourceRange.Copy targetCell

    MsgBox "数据已成功复制！", vbInformation
End Sub


Sub 单元格拷贝改进()
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetCell As Range

    ' 设置源范围和目标位置
'获取从 B2 开始的连续区域（即包含数据的矩形区域）
    Set sourceRange = Range("B2").CurrentRegion
'代码在执行过程中出错（例如 Sheet2 不存在），程序不会停止，而是继续执行下一行代码
    On Error Resume Next
'尝试设置目标工作表为当前工作簿
    Set targetSheet = ThisWorkbook.Sheets("Sheet2")
'关闭错误忽略模式，恢复默认的错误处理机制
    On Error GoTo 0
'判断 targetSheet 是否成功设置。如果目标工作表不存在，Set 操作失败，targetSheet 会被设置为 Nothing
    If targetSheet Is Nothing Then
        MsgBox "目标工作表不存在！", vbExclamation
        Exit Sub
    End If
    Set targetCell = targetSheet.Range("A1")
    
    ' 复制并粘贴
    sourceRange.Copy
    With targetCell
'确保列宽与源数据保持一致
        .PasteSpecial xlPasteColumnWidths
'粘贴源区域的全部内容，包括格式、公式、值
        .PasteSpecial xlPasteAll
    End With
    
    ' 清除复制模式
    Application.CutCopyMode = False
    MsgBox "数据复制完成！", vbInformation
End Sub


计算并显示两个选定区域的交集区域

'用于返回两个范围的交集
Sub Intersect交叉区域(r1 As Range, r2 As Range)
    Dim 单元格对象 As Range
    On Error Resume Next
    Set 单元格对象 = Excel.Application.Intersect(r1, r2)
    On Error GoTo 0
    If 单元格对象 Is Nothing Then
        MsgBox "不存在交叉区域", vbExclamation
    Else
        MsgBox "交叉区域地址为：" & 单元格对象.Address, vbInformation
    End If
    Set 单元格对象 = Nothing
End Sub

Sub 调用()
    Dim r1 As Range, r2 As Range
    On Error Resume Next
    Set r1 = Application.InputBox("请选择第一个区域：", Type:=8)
    Set r2 = Application.InputBox("请选择第二个区域：", Type:=8)
    On Error GoTo 0
    If r1 Is Nothing Or r2 Is Nothing Then
        MsgBox "未选择有效区域！", vbExclamation
        Exit Sub
    End If
    Call Intersect交叉区域(r1, r2)
End Sub


Sub 多选单元格()
    ' 声明变量
    Dim r1 As Range, r2 As Range
    Dim ws As Worksheet

    ' 指定工作表（可以修改为目标工作表）
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 遍历指定区域
    For Each r1 In ws.Range("A1:G7")
        If r1.Value = "孙兴华" Then
            ' 如果r2为空，初始化r2
            If r2 Is Nothing Then
                Set r2 = r1
            Else
                ' 合并新的单元格到r2
                Set r2 = Union(r2, r1)
            End If
        End If
    Next r1

    ' 如果找到符合条件的单元格
    If Not r2 Is Nothing Then
        r2.Interior.Color = vbRed ' 设置背景色为红色
        r2.Select ' 选中单元格区域
    Else
        MsgBox "未找到符合条件的单元格！", vbExclamation
    End If

    ' 清理对象
    Set r1 = Nothing
    Set r2 = Nothing
    Set ws = Nothing
End Sub


Sub 添加边框()
 Dim r1 As Range
 Set r1 = Range("A1:G7")
 With r1.Borders
 '边框线条样式
 .LineStyle = xlContinuous
 '边框线条粗细
 .Weight = xlThin
 '边框线条颜色
 .ColorIndex = 5
 End With
 '使用BorderAround方法为单元格区域添加一个加粗外框
 r1.BorderAround xlContinuous, xlMedium, 5
 Set r1 = Nothing
End Sub


在指定的单元格区域 A1:G7 中，设置内部边框为实线和虚线的组合，同时为整个区域添加一个加粗外框

Sub 外实内虚()
 Dim r1 As Range
 Set r1 = Range("A1:G7")
 With r1.Borders(xlInsideHorizontal) ‘内部水平
 .LineStyle = xlDot
 .Weight = xlThin
 .ColorIndex = 5
 End With
 With r1.Borders(xlInsideVertical) ‘内部垂直
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = 5
 End With
 r1.BorderAround xlContinuous, xlMedium, 5
 Set r1 = Nothing
End Sub

'一次排序可同时指定多个关键字（Key1 和 Key2）
'每次调用 .Sort 都能处理多个关键字，减少代码运行的次数
With Range("A1")
    .Sort Key1:="英语", order1:=xlDescending, _
          Key2:="语文", order2:=xlDescending, Header:=xlYes
    .Sort Key1:="数学", order1:=xlDescending, _
          Key2:="总成绩", order2:=xlDescending, Header:=xlYes
End With


根据 E2:E6 区域中定义的自定义排序顺序，对指定的表格数据按“部门”列进行排序

Sub SortByLists()
 Dim arr, 序号
 arr = Range("E2:E6")
 '通过AddCustomList方法为数组添加自定义序列
 Excel.Application.AddCustomList arr
 '返回数组在自定义序列中的序列号，保存在序号这个变量中
 序号 = Application.GetCustomListNum(arr)
 '因为OrderCustom从1开始，如果有一行表头我们就要加1
'表格包含标题行，标题不会参与排序
'指定按照自定义排序顺序排列
 Range("A1").Sort Key1:="部门", Order1:=xlAscending, Header:=xlYes, OrderCustom:=序号 + 1
 '使用DeleteCustomList删除新添加的自定义序列
 Application.DeleteCustomList 序号
End Sub


从某张表中筛选出当天生日的人员，并将这些人员的姓名复制到名为 "生日名单" 的工作表中，同时弹出消息框提示

Sub 提醒过生日的人名()
Dim i
i = 2: j = 1
Do While Range("A" & i) <> ""
'检查当前行的列 C 是否是当天的日期, C 中的日期与当天日期的月份和天数相同，则判定为当天生日
 If Month(Range("C" & i)) = Month(Date) And Day(Range("C" & i)) = Day(Date) Then
 MsgBox "今天是" & Range("A" & i) & "的生日"
 Sheets("生日名单").Range("A" & j) = Range("A" & i)
 j = j + 1
 End If
 i = i + 1
Loop
End Sub


对表格数据进行筛选，筛选条件是第 2 列等于单元格 H2 的值，并将筛选结果复制到单元格 K1 开始的位置

Sub 筛选()
'以单元格 K1 为起始点的连续数据区域,清空该区域的所有内容和格式
Range("K1").CurrentRegion.Clear
'筛选这个区域内第2列，等于H2单元格的名字
Range("A1").CurrentRegion.AutoFilter field:=2, Criteria1:=Range("H2")
'将这个区域复制到K1单元格
Range("A1").CurrentRegion.Copy Range("K1")
'取消自动筛选
Range("A1").CurrentRegion.AutoFilter
End Sub


为了避免在 筛选 过程中触发 Worksheet_Change 事件自身，从而避免递归调用或其他不必要的操作
'每当工作表中的单元格内容被修改时，Worksheet_Change 事件就会触发
'Target 是一个 Range 对象，表示发生变化的单元格区域
Private Sub Worksheet_Change(ByVal Target As Range)
    '关闭事件,将事件处理功能禁用，防止由于 筛选 操作改变单元格内容时再次触发 Worksheet_Change 事件，避免发生递归调用
    Excel.Application.EnableEvents = False
    '因为筛选本身就是修改单元格
    Call 筛选
    '打开事件
    Excel.Application.EnableEvents = True
End Sub
为什么使用 EnableEvents = False？
在 VBA 中，很多操作（如筛选、排序、复制等）会修改单元格内容或其他工作表属性。当这些修改发生时，会触发 Worksheet_Change 等事件。如果不禁用事件，可能会引发递归调用，也就是修改一个单元格内容后，事件再次触发，进而再修改另一个单元格，导致系统陷入无限循环，甚至导致程序崩溃。
禁用事件处理（EnableEvents = False）是防止这种递归调用的常见做法。通过这种方式，在进行某些操作（如筛选）时，可以避免触发不必要的事件。


Worksheet_Change 事件处理程序，用于处理特定单元格的更改操作
主要涉及以下两种情况：
当在第 2 列（假定为商品名称）输入数据时，根据输入的商品名称查找对应的商品信息（如价格、库存、销售数量等），并填充相关的单元格。
当在第 5 列（假定为销售数量）输入数据时，自动计算销售总额，并将结果显示在相邻单元格

Private Sub Worksheet_Change(ByVal Target As Range)
    '同时更改多个单元格时结束执行程序，CountLarge和count功能一样
    'CountLarge不会溢出，但是count会，xlsx单元格太多了，容易发生数据类型溢出
'如果被更改的单元格不止一个（例如，用户进行了批量粘贴），则不执行后续代码
    If Target.CountLarge <> 1 Then Exit Sub
    '输入数据为空时退出
    If Target.Value = "" Then Exit Sub
    '输入行号是第1行时退出,避免影响表头行
    If Target.Row = 1 Then Exit Sub
    Dim i
    '当输入等于第2列时
    If Target.Column = 2 Then
        '如果出错了，就是没找到，跳转到标签a
        On Error GoTo a
        'Match返回输入的商品名称来自参照表第几行,UCase 用于将输入转换为大写
        i = Excel.Application.WorksheetFunction.Match(UCase(Target.Value), Range("H:H"), 0)
        '禁止事件，防止将字母改为商品名称时，再次执行程序
        Excel.Application.EnableEvents = False
        With Target
            .Value = Range("I" & i).Value  '填充商品名称对应的值（如价格等）
            .Offset(0, -1).Value = Now  '记录当前时间
            .Offset(0, 1) = Range("J" & i).Value  '填充相关商品信息（如库存）
            .Offset(0, 2) = Range("K" & i).Value  '填充相关商品信息（如销售数量）
            '输入商品名称后，选中销售数量的单元格
            .Offset(0, 3).Select  '将焦点移动到销售数量的单元格
        End With
        Excel.Application.EnableEvents = True
        Exit Sub
a: MsgBox "没有该商品，请联系维护人员"
'在当前单元格的基础上，通过 Offset 方法填充相邻的单元格
        Target.Value = ""  '清空输入的商品名称
    Else
'如果修改的是第 5 列（假设为销售数量列），程序会根据销售数量和单价计算销售总额。
        If Target.Column = 5 Then
            Application.EnableEvents = False
'计算销售总额，并将结果放入销售总额所在的单元格。Offset(0, -1) 表示取得销售数量单元格左侧的单价。
            Target.Offset(0, 1) = Target * Target.Offset(0, -1)  '计算销售总额
            Cells(Target.Row + 1, 2).Select  '选中下一行的商品名称单元格
            Application.EnableEvents = True
        End If
    End If
End Sub


记录单元格内容的修改历史，并将修改记录保存为批注

'记录被选中单元格的内容，并将其存储为变量 r1。这个变量将用于后续的 Worksheet_Change 事件中，来与修改后的内容进行比较
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count <> 1 Then Exit Sub '选中多个单元格时退出程序
    If Target.Formula = "" Then '根据选中单元格中保存的数据，确定给变量r1赋什么值
        r1 = "空"
    Else
        r1 = Target.Text
    End If
End Sub

'记录被修改单元格的内容，并将修改记录添加为批注。批注内容包括修改前的内容、修改后的内容以及修改的时间。
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count <> 1 Then Exit Sub
    '定义变量保存单元格修改后的内容
    Dim r2
    '判断单元格是否被修改为空单元格
    If Target.Formula = "" Then
        r2 = "空"
    Else
        r2 = Target.Formula
    End If
    '如果单元格修改前后的内容一样则退出程序
    If r1 = r2 Then Exit Sub
    '定义一个批注变量
    Dim r3
    '定义一个变量保存批注内容
    Dim r4
    '将被修改单元格的批注赋给变量r3
'获取目标单元格的批注对象。如果单元格没有批注，则使用 Target.AddComment 创建一个新的批注
    Set r3 = Target.Comment
    '如果单元格中没有批注则新建批注
    If r3 Is Nothing Then Target.AddComment
    '将批注的内容保存到变量r4中
    r4 = Target.Comment.Text
    '重新修改批注的内容=原批注内容+当前日期和时间+原内容+修改后的新内容
    Target.Comment.Text Text:=r4 & Chr(10) & Format(Now(), "yyyy-mm-dd hh:mm") & "原内容:" & r1 & "修改为:" & r2
    '根据批注内容自动调整批注大小
    Target.Comment.Shape.TextFrame.AutoSize = True
End Sub


每隔 10 秒自动保存一次当前工作簿

Sub otime()
 '10秒后自动运行WbSave过程
' Now() + TimeValue("00:00:10")当前时间加上 10 秒后的时间
'"WbSave": 指定在设置的时间触发的过程
 Application.OnTime Now() + TimeValue("00:00:10"), "WbSave"
End Sub
Sub WbSave()
 ThisWorkbook.Save '保存本工作簿
 Call otime '再次运行otime过程
End Sub

打开工作簿立刻运行自动保存，通过输入框让用户设置时间间隔

为了让工作簿打开就自动运行：
Private Sub Workbook_Open()
Call otime
End Sub
Sub otime()
    Dim interval As String
    On Error Resume Next ' 防止用户输入错误导致代码中断
    interval = InputBox("请输入保存间隔时间 (格式：hh:mm:ss)", "设置间隔", "00:00:10")
    On Error GoTo 0
    ' 检查用户是否取消输入或输入为空
    If interval = "" Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    ' 验证时间格式
'使用 IsDate 检测用户输入的时间格式是否有效
    If IsDate(interval) Then
        Application.OnTime Now() + TimeValue(interval), "WbSave"
    Else
        MsgBox "输入的时间格式无效，请输入正确的时间格式 (如：00:00:10)。", vbExclamation
    End If
End Sub

Sub WbSave()
    On Error Resume Next ' 防止保存失败导致错误
    ThisWorkbook.Save ' 保存本工作簿
    On Error GoTo 0
    Call otime ' 再次运行otime过程
End Sub


Sub a()
Dim i, w1, arr
Set w2 = ActiveWorkbook
Set s2 = ActiveSheet
'调用 GetOpenFilename 显示文件选择对话框
'限制只能选择扩展名为 .xls 或 .xlsx 的文件
arr = Excel.Application.GetOpenFilename("Excel文件,*.xls*", MultiSelect:=True)
If IsArray(arr) Then
 For i = LBound(arr) To UBound(arr)
 Set w1 = Workbooks.Open(arr(i))
 For Each s1 In w1.Sheets
 s1.Copy after:=w2.Sheets(w2.Sheets.Count)
'使用 Split 将文件名按 . 分割，取分割后的第一个部分（即文件名）
 w2.Sheets(w2.Sheets.Count).Name = Split(w1.Name, ".")(0) & s1.Name
w1.Close
 Next
 Next
End If
End Sub

两个单选框控件 (xb1 和 xb2) 的交互功能，用户可以通过点击单选框在单元格 D1 中显示性别，并确保只有一个单选框被选中

Private Sub UserForm_Initialize()
    ' 设置初始化状态
    xb1.Value = True
    Range("D1").Value = "男"
End Sub

'用于更新性别的值，并实现单选框的互斥选择
Private Sub UpdateGender(selectedBox As Object, otherBox As Object, gender As String)
    ' 更新性别逻辑
'If selectedBox.Value = True Then检查当前选择的单选框是否被选中
    If selectedBox.Value = True Then
        Range("D1").Value = gender
        otherBox.Value = False
        MsgBox "当前选择：" & gender, vbInformation, "性别选择"
    End If
End Sub

Private Sub xb1_Click()
    ' 男单选框的点击事件
    UpdateGender xb1, xb2, "男"
End Sub

Private Sub xb2_Click()
    ' 女单选框的点击事件
    UpdateGender xb2, xb1, "女"
End Sub


用户通过输入原密码、新密码以及确认密码来修改密码。如果输入的原密码正确且新密码输入不为空且两次输入的密码相同，则修改密码并保存工作簿
Private Sub 更改密码_Click()
Dim 原密码
‘获取工作簿中名为 "用户密码" 的命名区域的值，即存储当前密码的名称。’
k = Names("用户密码").RefersTo
原密码 = InputBox("请输入原密码：", "提示")
'判断原密码是否正确
'Ctr是假定存在的一个函数，可能是对 原密码 进行某种验证或加密处理的函数
If Ctr(原密码) <> Evaluate(k) Then
 MsgBox "原密码输入错误，不能修改！", vbCritical, "错误"
 Exit Sub
End If
新密码1 = InputBox("请输入新密码:", "提示")
'判断新密码是否为空
If 新密码1 = "" Then
 MsgBox "新密码不能为空，修改没有完成！", vbCritical, "错误"
 Exit Sub
End If
新密码2 = InputBox("请再次输入新密码：", "提示")
'判断两次输入的亲密码是否相同
If 新密码1 = 新密码2 Then
 k = "=" & 新密码1
 ThisWorkbook.Save
 MsgBox "密码修改完成，下次登录请使用新密码！", vbInformation, "提示"
Else
 MsgBox "两次输入的密码不一致，修改没有完成！", vbCritical, "错误"
End If
End Sub

如何使用正则表达式（RegExp）在 Excel 中进行字符串的匹配、提取和替换

Sub RegExpDemoSyntax()
 Dim 正则, 结果集合, 结果
 字符串 = Range("A2").Value
 '给对象指定正则表达式对象
 'CreateObject函数用于创建各种外部对象，对象的完整名称就是参数
 Set 正则 = CreateObject("vbscript.regexp")
 'Pattern后面写正则表达式
'匹配以 Name: 开头的字符串，后面接任意字符
'再匹配以 Phone: 开头的字符串，后面接数字
 正则.Pattern = "Name:(.*?),Phone:(\d+)"
 'Global值为True返回所有符合要求的结果，反之只返回第一个符合要求的结果
 正则.Global = True
 'Execute(字符串)
 Set 结果集合 = 正则.Execute(字符串)
 If 结果集合.Count > 0 Then
 i = 2
 For Each 结果 In 结果集合
 Range("B" & i) = 结果.submatches(0)
 Range("C" & i) = 结果.submatches(1)
 Range("D" & i) = 正则.Replace(字符串, "$1$2")
 i = i + 1
 Next
 End If
set 结果集合 = Nothing
 Set 正则 = Nothing
End Sub


通过正则表达式对公式进行解析和替换，主要用于在 Excel 中处理带有表名和单元格引用的公式

Sub RegExpDemoReplace()
 Dim 正则, i
 '使用后期绑定正则对象
 Set 正则 = CreateObject("vbscript.regexp")
 '正则表达式
'非 ! 的任意字符
'匹配 1 到 3 个大写字母（列标），后跟 1 到 6 个数字（行号）
 正则.Pattern = "[^!]([A-Z]{1,3}\d{1,6})"
 正则.Global = True
 i = 1
 Do While Range("A" & i) <> ""
 '在公式前添加一个空格，确保第一个单元格引用可以被正则匹配成功
 公式 = " " & Range("B" & i)
 表名 = Range("A" & i) & "!"
 Set 结果集合 = 正则.Execute(公式)
 '成功匹配的对象（字符组）数目
 If 结果集合.Count > 0 Then
 For j = 1 To 结果集合.Count
'提取捕获组中的内容，即匹配到的单元格引用
 单元格 = 结果集合(j - 1).submatches(0)
 公式 = Replace(公式, 单元格, 表名 & 单元格)
 Next
Range("C" & i) = Trim(公式)
 End If
 i = i + 1
 Loop
 Set 正则 = Nothing
 Set 结果集合 = Nothing
End Sub


单元格内容中移除以 * 结尾的数字,对清理后的字符串作为公式进行计算,将计算结果存储在当前行右侧的列中
Sub RegExpDemo()
 Dim 正则
 '使用后期绑定创建正则对象
 Set 正则 = CreateObject("vbscript.regexp")
 '指定正则匹配字符串
'匹配以 * 结尾的数字
 正则.Pattern = "\d+\*"
 '设置为全局搜索模式
 正则.Global = True
'遍历列 B 中的非空单元格，从 B2 开始到列 B 的最后一行
'定位列 B 的最后一个非空单元格
 For Each 单元格 In Range([B2], Cells(Rows.Count, "B").End(xlUp))
'去除单元格值前后空格
'将匹配到的以 * 结尾的数字替换为空字符串（即删除）
 新字符串 = 正则.Replace(Trim(单元格.Value), "")
'计算清理后的字符串作为公式的结果,将结果写入当前单元格右侧的单元格
 单元格.Offset(0, 1).Value = Application.Evaluate(新
字符串)
Next
 Set 正则 = Nothing
End Sub


提取日期和金额

Sub ExtractDateAndAmount()
    Dim 正则 As Object, i As Long
    Dim 结果集合 As Object, 字符串 As String

    ' 初始化正则表达式对象
    Set 正则 = CreateObject("vbscript.regexp")
'匹配日期格式,匹配金额格式
    正则.Pattern = "(\d{4}-\d{2}-\d{2}|\d{4}.\d{2}.\d{2}).*?(([A-Z]{3})*\d+[\d.,]*元)"
    正则.Global = True

    i = 2
    Do While Range("A" & i) <> ""
        字符串 = Range("A" & i).Value
'对当前单元格的字符串进行匹配，并将结果存储到 结果集合 中
        Set 结果集合 = 正则.Execute(字符串)
        
        If 结果集合.Count > 0 Then
            Range("B" & i) = 结果集合(0).submatches(0)
            Range("C" & i) = 结果集合(0).submatches(1)
        Else
            Range("B" & i) = "无匹配"
            Range("C" & i) = "无匹配"
        End If
        
        i = i + 1
    Loop

    ' 释放对象
    Set 正则 = Nothing
    Set 结果集合 = Nothing
End Sub



Sub ExtractChineseCharacters()
    Dim 正则 As Object
    Dim 单元格 As Range
    Dim 字符串 As String

    ' 初始化正则表达式对象
    Set 正则 = CreateObject("vbscript.regexp")
'[^...] 表示取反，即匹配 非汉字字符
    正则.Pattern = "[^一-龥]" ' 匹配非汉字字符
    正则.Global = True

    ' 遍历列 A 中的所有单元格
    For Each 单元格 In Range([A2], Cells(Rows.Count, 1).End(xlUp))
'去除单元格值两端的空格
        字符串 = Trim(单元格.Value)
'替换 字符串 中的所有非汉字字符为空字符串，即提取出纯汉字
        单元格.Offset(0, 1).Value = 正则.Replace(字符串, "") ' 提取汉字并写入列 B
    Next

    ' 释放对象
    Set 正则 = Nothing
End Sub


从一个字典对象中读取数据并填充到指定的查询结果区域
Sub 从字典中读取数据()
Dim 字典, arr, brr
Set 字典 = CreateObject("Scripting.Dictionary")
arr = Range("A1").CurrentRegion
'对字典里的每个key赋值
For i = 2 To UBound(arr)
'将主数据区域中的第一列作为键，第二列作为值存入字典
 字典(arr(i, 1)) = arr(i, 2)
Next
'查询区域
brr = Range("D1:E" & Cells(Rows.Count, 
"D").End(xlUp).Row)
For i = 2 To UBound(brr)
 '如果字典存在Key值
'判断查询区域中的值是否存在于字典的键中
 If 字典.Exists(brr(i, 1)) Then
 '获取人名对应的条目，分数那列等于人名那列的item
 brr(i, 2) = 字典(brr(i, 1))
 Else
 brr(i, 2) = "查无此人"
 End If
Next
Range("D1:E" & Cells(Rows.Count, "D").End(xlUp).Row) 
= brr
Set 字典 = Nothing
End Sub
