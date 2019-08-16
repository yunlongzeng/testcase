Imports Microsoft.Office.Tools.Ribbon
Imports System.Data.OleDb
Imports System.Data

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
    '设置用例目录所在的列数
    Public col1 = 1
    '设置用例名称所在的列数
    Public col2 = 2
    '设置需求ID所在的列数
    Public col3 = 3
    '设置前置条件所在的列数
    Public col4 = 4
    '设置用例步骤所在的列数
    Public col5 = 5
    '设置预期结果所在的列数
    Public col6 = 6
    '设置用例类型所在的列数
    Public col7 = 7
    '设置用例状态所在的列数
    Public col8 = 8
    '设置用例等级所在的列数
    Public col9 = 9
    '设置创建人所在的列数
    Public col10 = 10
    '初始化用例标题行数
    Public text_row = 2
    '设置创建人名字
    Public create_name = "曾云龙"

    '项目相关代码
    '第几行开始
    Public row = 3
    '第几列开始
    Public col = 3
    '设置完成情况所在的列数
    Public p_col1 = 10
    '设置预计开始时间所在的列数
    Public p_col2 = 5
    '设置预计结束时间所在的列数
    Public p_col3 = 6
    '设置状态所在的列数
    Public p_col4 = 3
    '设置实际开始所在的列数
    Public p_col5 = 8
    '设置实际结束所在的列数
    Public p_col6 = 9
    '设置备注所在的列数
    Public p_col7 = 7

    '统计工时
    Public row_work = 11
    Public col_work = 1
    '一天到七天的颜色号
    Public color_1 = 20
    Public color_2 = 35
    Public color_3 = 19
    Public color_4 = 8
    Public color_5 = 6
    Public color_6 = 4
    Public color_7 = 3

    '设置路径
    Public path_address = "C:\Users\Public\Documents"
    '设置文件保存路径
    Public filename_path = "C:\Users\12959\测试资料\xmind\"
    '设置excel路径
    Public path = path_address + "\testcase.xlsx"
    '设置database地址
    Public data_address = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_address + "\testcase.accdb"
    '设置项目管理工作excel路径
    Public work_path = "C:\Users\12959\测试资料\项目管理\work.xlsx"
    Public Function write_title(cl1, cl2, cl3, cl4, cl5, cl6, cl7, cl8, cl9, cl10， xlSheet)   '设置首行title    

        xlSheet.Cells(1, cl1).value = "用例目录"
        xlSheet.Cells(1, cl2).value = "用例名称"
        xlSheet.Cells(1, cl3).value = "需求ID"
        xlSheet.Cells(1, cl4).value = "前置条件"
        xlSheet.Cells(1, cl5).value = "用例步骤"
        xlSheet.Cells(1, cl6).value = "预期结果"
        xlSheet.Cells(1, cl7).value = "用例类型"
        xlSheet.Cells(1, cl8).value = "用例状态"
        xlSheet.Cells(1, cl9).value = "用例等级"
        xlSheet.Cells(1, cl10).value = "创建人"

    End Function

    Public Function write_title1(cl1, cl2, cl3, cl4, xlSheet)   '设置首行title

        xlSheet.Cells(1, cl1).value = "完成情况"
        xlSheet.Cells(1, cl2).value = "实际开始"
        xlSheet.Cells(1, cl3).value = "实际结束"
        xlSheet.Cells(1, cl4).value = "备注"

    End Function
    Public Function get_col_char(n)   '设置首行title

        Dim a = n \ 26
        Dim b = n Mod 26

        If (a > 0) Then
            Return CStr(Chr(a + 64)) + CStr(Chr(b + 65))
        End If

        Return CStr(Chr(b + 65))

    End Function
    Public Function check_copy() '检查是否将xmind复制到第一行第一列

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        If (CStr(xlSheet.Cells(1, 1).value) = "") Then
            Return "false"
        ElseIf (CStr(xlSheet.Cells(1, 1).value) = "用例目录") Then
            Return "already"
        Else
            Return "true"
        End If

    End Function
    Public Function count_blank(row1, blank_front)
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim val = 0
        For i = 1 To 1000
            If (CStr(xlSheet.Cells(row1, i).value) = "") Then
                val += 1
                If val = blank_front + 2 Then
                    Return val
                End If
            Else
                Return val
            End If
        Next
    End Function
    Public Function get_text(row1, blank_after)
        '获得当前激活的sheet       
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '初始化有句号的列，0代表该行没有
        Dim col22 = 0
        For i = blank_after + 1 To 1000
            If (CStr(xlSheet.Cells(row1, i).value) = "") Then
                '如果为空则跳出该循环
                col22 = 0
                'MsgBox(col22)
                Return col22
            ElseIf InStr(CStr(xlSheet.Cells(row1, i).Value), "。") Then
                'ElseIf InStr(CStr(xlSheet.Cells(row1, i).Value), "。") Or InStr(CStr(xlSheet.Cells(row1, i).Value), ".") Then
                '获取有句号的列
                col22 = i
                'MsgBox(col22)
                Return col22
            End If
        Next
    End Function
    Public Function clear_text(a, blank_before, row_num, b)
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        'MsgBox(a)
        'MsgBox(blank_before)
        'MsgBox(row_num)
        'MsgBox(b)
        For i = 1 To a + 1
            If (CStr(xlSheet.Cells(i, 1).value) <> "") Then
                xlSheet.Cells(i, 1).value = ""
            End If
            If (CStr(xlSheet.Cells(i, 3).value) <> "") Then
                xlSheet.Cells(i, 3).value = ""
            End If
            If (CStr(xlSheet.Cells(i, 5).value) <> "") Then
                xlSheet.Cells(i, 5).value = ""
            End If
            If (CStr(xlSheet.Cells(i, 7).value) <> "") Then
                xlSheet.Cells(i, 7).value = ""
            End If
            If (CStr(xlSheet.Cells(i, 8).value) <> "") Then
                xlSheet.Cells(i, 8).value = ""
            End If
            If (CStr(xlSheet.Cells(i, 9).value) <> "") Then
                xlSheet.Cells(i, 9).value = ""
            End If

            If b > 10 Then

                For j = 11 To b
                    If (CStr(xlSheet.Cells(i, j).value) <> "") Then
                        xlSheet.Cells(i, j).value = ""
                    End If
                Next

            End If
        Next
        For i = a + 2 To row_num
            For j = 1 To b
                If (CStr(xlSheet.Cells(i, j).value) <> "") Then
                    xlSheet.Cells(i, j).value = ""
                End If
            Next
        Next

    End Function
    Public Function getLineNumber(xlsheet, col, row)  '获取excel行数函数

        Dim a = row
        For k = row To 1000
            If CStr(xlsheet.Cells(k, col).value) = "" Then
                a = k - 2
                Return a
            End If
        Next

    End Function
    Public Function get_sqltext(text, ver)

        Dim arr() As String
        Dim arr1(2) As String
        If InStr(text, "，") Then
            arr = Split(text, "，")
        ElseIf InStr(text, ",") Then
            arr = Split(text, ",")
        End If
        If ver = "PC" Then
            If InStr(arr(0), "】") Then
                arr1 = Split(arr(0), "】")
            ElseIf InStr(arr(0), "]") Then
                arr1 = Split(arr(0), "]")
            Else
                MsgBox("Please use format [PC]")
                arr1(1) = "0"
            End If
        ElseIf ver = "APP" Then
            If InStr(arr(0), "】") Then
                arr1 = Split(arr(0), "】")
            ElseIf InStr(arr(0), "]") Then
                arr1 = Split(arr(0), "]")
            Else
                MsgBox("Please use format [APP]")
                arr1(1) = "0"
            End If
        Else
            'MsgBox(arr(0))
            'MsgBox(arr(1))
            If InStr(arr(0), "H5") Or InStr(arr(0), "小程序") Or InStr(arr(0), "h5") Then
                If InStr(arr(0), "】") Then
                    arr1 = Split(arr(0), "】")
                ElseIf InStr(arr(0), "]") Then
                    arr1 = Split(arr(0), "]")
                Else
                    MsgBox("Please use format [H5  小程序]")
                    arr1(1) = "0"
                End If
            End If
        End If
        Return "select step from " + ver + " where title = '" + Trim(arr1(1)) + "'"

    End Function
    Public Function searchdata(sqltxt, a)
        'MsgBox("1")
        Dim Con As New OleDbConnection(data_address)
        Con.Open()
        Dim str1(a - 1) As String
        For i = 0 To a - 1
            'MsgBox("2")
            Dim strData As String
            strData = String.Empty
            Dim objCommand As New OleDbCommand(sqltxt(i), Con)
            Dim objReader As OleDbDataReader
            objReader = objCommand.ExecuteReader()
            'MsgBox("3")
            While objReader.Read()
                For intindex As Integer = 0 To objReader.FieldCount - 1
                    strData &= objReader.Item(intindex).ToString
                Next
            End While
            'MsgBox(strData)

            str1(i) = ""
            If strData <> "" Then
                str1(i) = strData
                'MsgBox("4")
            Else
                str1(i) = ""
            End If
            'MsgBox("5")
        Next
        Con.Close()
        Return str1
    End Function
    Public Function openFile(path)
        If System.IO.File.Exists(path) = False Then
            MsgBox("The data doesn't exist!")
        ElseIf System.IO.File.Exists(path) Then
            Dim xlApp As Excel.Application      '定义 Excel 程序
            Dim xlBook As Excel.Workbook      '定义 Excel 工作簿
            Dim xlSheet As Excel.Worksheet    '定义 Excel 工作表

            '3、进行Excel操作
            xlApp = CreateObject("Excel.Application") '创建EXCEL对象
            xlBook = xlApp.Workbooks.Open(path) '打开已经存在的EXCEL工件簿文件       
            xlApp.Visible = True 'Excel的可见性
            xlSheet = xlBook.Worksheets(1) '设置活动工作表 表名可用 1\2\3\4代替
        End If
    End Function
    Public Function updatedata(sqltxt)

        Dim conn As New OleDb.OleDbConnection(data_address)
        conn.Open() '打开连接
        Dim da As New OleDb.OleDbDataAdapter()
        da.SelectCommand = New OleDbCommand(sqltxt, conn)
        da.SelectCommand.ExecuteNonQuery()

        conn.Close()

    End Function
    Public Function updateaccess(sqltxt, a)
        'MsgBox("1")
        Dim Con As New OleDbConnection(data_address)
        Con.Open()
        For i = 0 To a - 1
            'MsgBox(i)
            Dim strData As String
            strData = String.Empty
            Dim objCommand As New OleDbCommand(sqltxt(i, 0), Con)
            Dim objReader As OleDbDataReader
            objReader = objCommand.ExecuteReader()
            'MsgBox("3")
            While objReader.Read()
                For intindex As Integer = 0 To objReader.FieldCount - 1
                    strData &= objReader.Item(intindex).ToString
                Next
            End While
            'MsgBox(strData)

            If strData = "" Then
                updatedata(sqltxt(i, 1))
                'MsgBox(sqltxt(i, 1))
            End If
            'MsgBox("5")
        Next
        Con.Close()
    End Function
    Public Function copy_to_access(xlsheet1, arr1, a1, ver)
        If a1 <> 0 Then
            For i = 0 To a1 - 1
                arr1(i, 0) = "select step from " + ver + " where title = '" + Trim(xlsheet1.Cells(i + 2, 2).value) + "'"
                arr1(i, 1) = "insert into " + ver + "(title,step) values('" + xlsheet1.Cells(i + 2, 2).value + "','" + xlsheet1.Cells(i + 2, 3).value + "')"
            Next
            updateaccess(arr1, a1)
        End If
    End Function

    Public Function get_line(xlsheet, a, b)
        For k = a To 1000
            If CStr(xlsheet.Cells(k, b).value) = "" Then
                a = k - a
                Return a
            End If
        Next
    End Function

    Public Function merge_cell(xlSheet, r, c, text)
        '合并单元格
        xlSheet.Cells(r, c).value = text
        Dim str1 = get_col_char(c - 1)
        Dim str2 = str1.Trim() + Str(r).Trim()
        Dim str3 = str1.Trim() + Str(r + 1).Trim()
        Dim str4 = str2.Trim() & ":".Trim() & str3.Trim()
        xlSheet.Range(str4).MergeCells = True
    End Function
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        '将xmind的内容以句号分割，前面那句代表标题，后面那句代表预期结果

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '检查是否将xmind复制到第一行第一列,复制正确true返回
        Dim flag = check_copy()
        '初始化前一行空格数
        Dim blank_front = 0
        '初始化后一行空格数
        Dim blank_after = 0
        Dim row_num = 0
        '获取excel行数
        Dim a = 0
        '获取最大列数
        Dim b = 0
        Dim text_row_re = text_row
        '获取带句号的行数
        Dim full_stop_num = 0
        '获取xmind的总分支
        Dim xmind_num = 0

        If flag = "true" Then
            '文件拷贝正确就执行执行筛查操作


            '设置列宽
            xlSheet.Columns(col2).ColumnWidth = 60
            xlSheet.Columns(col4).ColumnWidth = 30
            xlSheet.Columns(col5).ColumnWidth = 40
            xlSheet.Columns(col6).ColumnWidth = 90

            '自动换行
            xlSheet.Columns(col2).WrapText = True
            xlSheet.Columns(col4).WrapText = True
            xlSheet.Columns(col5).WrapText = True
            xlSheet.Columns(col6).WrapText = True

            For i = 1 To 1000
                '创建一个循环，退出循环的条件是后一行空的累计等于上一行空行加2
                blank_after = count_blank(i, blank_front)
                If blank_after <= blank_front Then  '计算xmind总共多少分支
                    xmind_num += 1
                End If
                For k = blank_after + 1 To 1000
                    If CStr(xlSheet.Cells(i, k).value) = "" Then
                        If b < k Then
                            b = k
                        End If
                        Exit For
                    End If
                Next

                If blank_after = blank_front + 2 Then
                    row_num = i
                    Exit For
                Else
                    blank_front = blank_after

                    '获得哪一列包含句号
                    Dim col_text = get_text(i, blank_after)
                    If col_text <> 0 Then
                        a += 1
                        xlSheet.Cells(a + 1, 10).value = create_name
                        'MsgBox(xlSheet.Cells(i, col_text).value)
                        Dim arr() As String
                        If InStr(CStr(xlSheet.Cells(i, col_text).Value), "。") Then
                            arr = Split(xlSheet.Cells(i, col_text).value, "。")
                            full_stop_num += 1
                            'ElseIf InStr(CStr(xlSheet.Cells(i, col_text).Value), ".") Then
                            'arr = Split(xlSheet.Cells(i, col_text).value, ".")
                            'full_stop_num += 1
                        End If

                        'MsgBox(xlSheet.Cells(i, col_text).value)
                        '将内容填写到用例名称栏

                        xlSheet.Cells(text_row_re, col2).value = arr(0)
                        'MsgBox(arr(0))
                        '将内容填写到预期结果栏
                        xlSheet.Cells(text_row_re, col6).value = arr(1)

                        '将内容填写到前置条件
                        If (CStr(xlSheet.Cells(i + 1, col_text + 1).value) = "") Then
                            xlSheet.Cells(text_row_re, col4).value = "无"
                        Else
                            xlSheet.Cells(text_row_re, col4).value = xlSheet.Cells(i + 1, col_text + 1).value
                        End If

                        text_row_re += 1
                    End If
                End If


            Next

            '查询xmind内容是否包含换行
            For i = row_num To row_num + 4
                If (CStr(xlSheet.Cells(i + 1, 1).value) <> "") Then
                    MsgBox("Line " & i + 1 & " is blank!")
                End If

            Next

            '清空copy进来的内容
            If a <> 0 Then

                clear_text(a, blank_front, row_num, b)

            End If


            '设置首行title
            write_title(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, xlSheet)


            '隐藏用例目录、需求ID
            xlSheet.Columns(col1).Hidden = True
            xlSheet.Columns(col3).Hidden = True

            xlSheet.Cells(1, 1).Interior.ColorIndex = 0  '设置单元格背景颜色
            If full_stop_num <> xmind_num Then
                MsgBox("The lines total " & xmind_num & ", testcases are only " & full_stop_num & " ,please check!")
            Else
                MsgBox("Sorted!")
            End If
        ElseIf flag = "false" Then
            MsgBox("Please copy to the first cell!")
            xlSheet.Cells(1, 1).Interior.ColorIndex = 27  '设置单元格背景颜色
        Else
            MsgBox("Already sorted!")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim a = getLineNumber(xlSheet, col2, 2)
        'MsgBox(a)
        '判断是否用例名称包含】
        Dim flag_title = True

        Dim sql_text(a) As String   '初始化sql语句
        For i = 2 To a + 1

            If InStr(CStr(xlSheet.Cells(i, col2).value), "】") Or InStr(CStr(xlSheet.Cells(i, col2).value), "]") Then
                If InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "[pc]") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "【pc】") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "[pc】") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "【pc]") Then
                    sql_text(i - 2) = get_sqltext(CStr(xlSheet.Cells(i, col2).value), "PC")
                    'MsgBox(sql_text(i - 2))
                ElseIf InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "[app]") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "【app】") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "[app】") Or InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "【app]") Then
                    sql_text(i - 2) = get_sqltext(CStr(xlSheet.Cells(i, col2).value), "APP")
                    'MsgBox(sql_text(i - 2))
                Else
                    sql_text(i - 2) = get_sqltext(CStr(xlSheet.Cells(i, col2).value), "H5")
                    'MsgBox(sql_text(i - 2))
                End If
            Else
                flag_title = False
            End If
        Next

        If flag_title = True Then
            '数据库查询
            Dim arr2 = searchdata(sql_text, a)

            For i = 0 To a - 1
                xlSheet.Cells(i + 2, col5).value = arr2(i)
            Next

        Else
            MsgBox("The title doesn't exist [], please fill correct!")
            End If
        Dim r_num = 0
        For i = 2 To a + 2
            If CStr(xlSheet.Cells(i, 5).value) = "" Then
                Exit For
            End If
            r_num += 1
        Next

        If r_num = a Then
            MsgBox("Finished!")
        Else
            MsgBox(a - r_num & " lines lack of steps!")
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '取消隐藏用例目录、需求ID
        xlSheet.Columns(col1).Hidden = False
        xlSheet.Columns(col3).Hidden = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim filename As String '获取文件路径
        If Trim(EditBox1.Text) <> "" Then

            filename = filename_path + Trim(EditBox1.Text) + ".xlsx"
            Try
                xlWorkbook.SaveAs(Filename:=filename, FileFormat:=51)
            Catch ex As Exception
            End Try
        Else
            MsgBox("Please input filename!")
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        openFile(path)
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        '打开excel
        Dim xlApp As New Excel.Application
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlSheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlSheet2 As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlSheet3 As Microsoft.Office.Interop.Excel.Worksheet

        xlApp.Visible = False
        xlApp.Workbooks.Open(path)

        xlBook = xlApp.Workbooks(1)
        xlSheet1 = xlBook.Sheets("PC")
        xlSheet2 = xlBook.Sheets("H5-小程序")
        xlSheet3 = xlBook.Sheets("APP")

        '获取excel行数
        Dim a1 = getLineNumber(xlSheet1, 2, 2)
        Dim a2 = getLineNumber(xlSheet2, 2, 2)
        Dim a3 = getLineNumber(xlSheet3, 2, 2)

        Dim arr1(a1, 1) As String
        Dim arr2(a2, 1) As String
        Dim arr3(a3, 1) As String

        copy_to_access(xlSheet1, arr1, a1, "PC")
        copy_to_access(xlSheet2, arr2, a2, "H5")
        copy_to_access(xlSheet3, arr3, a3, "APP")

        xlApp.Quit()
        GC.Collect()
        MsgBox("Updated")

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim n = getLineNumber(xlSheet, col, row)

        If (CStr(xlSheet.Cells(row, 2).value) = "") Then
            MsgBox("No data!")
        ElseIf xlSheet.Range("A1:A2").MergeCells = True Then
            '分解单元格
            xlSheet.Range("A1:A2").MergeCells = False
            xlSheet.Range("B1:B2").MergeCells = False
            xlSheet.Range("C1:C2").MergeCells = False
            xlSheet.Range("D1:D2").MergeCells = False
            xlSheet.Range("E1:E2").MergeCells = False
            xlSheet.Range("F1:F2").MergeCells = False
            xlSheet.Range("G1:G2").MergeCells = False
            xlSheet.Range("H1:H2").MergeCells = False
            xlSheet.Range("I1:I2").MergeCells = False
            xlSheet.Range("J1:J2").MergeCells = False
            '删除第2行
            xlSheet.Cells(2, 1).EntireRow.Delete
            MsgBox("Please click again!")
        Else

            '清空没用的数据
            For i = row - 1 To n + row - 1
                xlSheet.Cells(i, p_col1).Interior.ColorIndex = 0  '设置单元格背景颜色
                xlSheet.Cells(i, p_col1).value = ""
                xlSheet.Cells(i, p_col5).value = ""
                xlSheet.Cells(i, p_col7).value = ""
            Next

            '插入第一行
            xlSheet.Cells(1, 1).EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)

            '合并单元格
            xlSheet.Range("A1:A2").MergeCells = True
            xlSheet.Range("B1:B2").MergeCells = True
            xlSheet.Range("C1:C2").MergeCells = True
            xlSheet.Range("D1:D2").MergeCells = True
            xlSheet.Range("E1:E2").MergeCells = True
            xlSheet.Range("F1:F2").MergeCells = True
            xlSheet.Range("G1:G2").MergeCells = True
            xlSheet.Range("H1:H2").MergeCells = True
            xlSheet.Range("I1:I2").MergeCells = True
            xlSheet.Range("J1:J2").MergeCells = True

            write_title1(p_col1, p_col5, p_col6, p_col7, xlSheet)

            For i = row To n + row - 1
                xlSheet.Cells(i, p_col2).value = Replace(xlSheet.Cells(i, p_col2).value, "-", "/")
                xlSheet.Cells(i, p_col3).value = Replace(xlSheet.Cells(i, p_col3).value, "-", "/")
                Dim str() = Split(xlSheet.Cells(i, p_col2).value, "/")
                Dim str1() = Split(xlSheet.Cells(i, p_col3).value, "/")
                If str(0) = "" Then
                    xlSheet.Cells(i, p_col1).value = "未填写日期"
                    xlSheet.Cells(i, p_col1).Interior.ColorIndex = 15  '设置单元格背景颜色
                Else
                    Dim fromDate1 = "#" + str(1) + "/" + str(2) + "/" + str(0) + "#"
                    Dim dif1 = DateDiff(“d”, fromDate1, Now)
                    Dim fromDate2 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"
                    Dim dif2 = DateDiff(“d”, fromDate2, Now)

                    '任务完成，自动填充实际开始结束时间
                    If (xlSheet.Cells(i, ).value = "已完成" Or xlSheet.Cells(i, p_col4).value = "已实现") Then
                        xlSheet.Cells(i, p_col1).value = "已完成"
                        xlSheet.Cells(i, p_col1).Interior.ColorIndex = 4  '设置单元格背景颜色
                        xlSheet.Cells(i, p_col5).value = xlSheet.Cells(i, p_col2).value
                        xlSheet.Cells(i, p_col6).value = xlSheet.Cells(i, p_col3).value

                    ElseIf (xlSheet.Cells(i, p_col4).value = "未开始" Or xlSheet.Cells(i, p_col4).value = "规划中") Then
                        If dif1 > 0 Then
                            xlSheet.Cells(i, p_col1).value = "开始日期已延期"
                            xlSheet.Cells(i, p_col1).Interior.ColorIndex = 26  '设置单元格背景颜色
                        End If
                        If dif1 = 0 Then
                            xlSheet.Cells(i, p_col1).value = "今日需开始"
                            xlSheet.Cells(i, p_col1).Interior.ColorIndex = 27  '设置单元格背景颜色
                        End If

                        '任务完成，自动填充实际开始时间
                    ElseIf (xlSheet.Cells(i, p_col4).value = "进行中" Or xlSheet.Cells(i, p_col4).value = "实现中") Then
                        xlSheet.Cells(i, p_col5).value = xlSheet.Cells(i, p_col2).value
                        If dif2 > 0 Then
                            xlSheet.Cells(i, p_col1).value = "结束日期已延期"
                            xlSheet.Cells(i, p_col1).Interior.ColorIndex = 3  '设置单元格背景颜色
                        End If
                    End If
                End If
            Next
            For i = 1 To 10
                xlSheet.Columns(i).AutoFit  '设置自适应列宽
            Next

            '生成autofilter
            xlSheet.Range("A1:J1").AutoFilter(Field:=1)
        End If

        Dim xl = Globals.ThisAddIn.Application
        xl.ActiveWindow.SplitColumn = 1
        xl.ActiveWindow.SplitRow = 0
        xl.Application.ActiveWindow.FreezePanes = True

    End Sub

    Private Sub work_Click(sender As Object, e As RibbonControlEventArgs) Handles work.Click
        openFile(work_path)
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        If xlSheet.Range("A1:A2").MergeCells = False Then
            MsgBox("Please click Sort button!")
        Else
            '初始化单元格格式
            xlSheet.Rows(1).NumberFormatLocal = "G/通用格式"  '@为文本格式

            '获取excel行数
            Dim n = getLineNumber(xlSheet, col, row)
            xlSheet.Cells(1, 14).value = "=MIN(E3:E" & n + 2 & ")"
            xlSheet.Cells(1, 15).value = "=MAX(F3:F" & n + 2 & ")"
            xlSheet.Cells(1, 16).value = "=O1-N1"
            xlSheet.Cells(1, 11).value = xlSheet.Cells(1, 14).value
            xlSheet.Cells(2, 11).value = "=TEXT(""" + xlSheet.Cells(1, 11).value + """," + """AAA""" + ")"
            xlSheet.Cells(2, 11).font.size = 8
            Dim day1 = xlSheet.Cells(1, 14).value
            Dim a = Integer.Parse(xlSheet.Cells(1, 16).value)

            For i = row To n + row - 1
                For j = 11 To a + 17
                    If (CStr(xlSheet.Cells(i, j).value) <> "") Then
                        xlSheet.Cells(i, j).value = ""
                    End If
                Next
            Next

            '自动填充日期
            For i = 12 To a + 17
                Dim day2 As DateTime = day1.AddDays(1)
                xlSheet.Cells(1, i).value = day2
                day1 = day2
                xlSheet.Cells(2, i).value = "=TEXT(""" + day2 + """," + """AAA""" + ")"
                xlSheet.Cells(2, i).font.size = 8
            Next

            For i = row To n + row - 1
                If (CStr(xlSheet.Cells(i, p_col2).value) = "-") Then
                    xlSheet.Cells(i, p_col2).value = Replace(xlSheet.Cells(i, p_col2).value, "-", "/")
                End If
                If (CStr(xlSheet.Cells(i, p_col3).value) = "-") Then
                    xlSheet.Cells(i, p_col3).value = Replace(xlSheet.Cells(i, p_col3).value, "-", "/")
                End If
                Dim str() = Split(xlSheet.Cells(i, p_col2).value, "/")
                Dim str1() = Split(xlSheet.Cells(i, p_col3).value, "/")
                If str(0) = "" Then
                    Continue For
                Else
                    For j = 11 To a + 17
                        Dim str2() = Split(xlSheet.Cells(1, j).value, "/")
                        Dim fromDate3 = "#" + str2(1) + "/" + str2(2) + "/" + str2(0) + "#"
                        Dim fromDate1 = "#" + str(1) + "/" + str(2) + "/" + str(0) + "#"
                        Dim dif1 = DateDiff(“d”, fromDate1, fromDate3)
                        Dim fromDate2 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"
                        Dim dif2 = DateDiff(“d”, fromDate2, fromDate3)
                        If dif1 >= 0 And dif2 <= 0 Then
                            xlSheet.Cells(i, j).value = "₪"
                        End If
                    Next
                End If
            Next
            For j = 11 To a + 17
                xlSheet.Columns(j).ColumnWidth = 1
            Next
        End If



    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim n = getLineNumber(xlSheet, col, row)

        Dim a = 0
        For i = 11 To 10000
            If (CStr(xlSheet.Cells(1, i).value) <> "") Then
                a += 1
            Else
                Exit For
            End If
        Next
        For i = row To n + row - 1
            For j = 11 To a + 10
                xlSheet.Cells(i, j).Interior.ColorIndex = 0  '设置单元格背景颜色
            Next
        Next
        For i = row To n + row - 1
            If (CStr(xlSheet.Cells(i, p_col5).value) <> "") Then
                Dim str1() = Split(xlSheet.Cells(i, p_col5).value, "/")
                Dim fromDate1 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"  '实际开始日期
                Dim str4() = Split(xlSheet.Cells(i, p_col3).value, "/")
                Dim fromDate4 = "#" + str4(1) + "/" + str4(2) + "/" + str4(0) + "#"  '计划结束日期
                For j = 11 To a + 10
                    Dim str2() = Split(xlSheet.Cells(1, j).value, "/")
                    Dim fromDate2 = "#" + str2(1) + "/" + str2(2) + "/" + str2(0) + "#"  '表格中的日期
                    If (CStr(xlSheet.Cells(i, 9).value) = "") Then
                        Dim dif1 = DateDiff(“d”, fromDate1, fromDate2)
                        Dim dif2 = DateDiff(“d”, fromDate2, Now)
                        Dim dif3 = DateDiff(“d”, fromDate4, fromDate2)
                        If dif1 >= 0 And dif2 > 0 Then
                            xlSheet.Cells(i, j).Interior.ColorIndex = 4  '设置单元格背景颜色
                            If dif3 > 0 Then
                                xlSheet.Cells(i, j).Interior.ColorIndex = 3  '设置单元格背景颜色
                            End If
                        End If
                    Else
                        Dim str3() = Split(xlSheet.Cells(i, 9).value, "/")
                        Dim fromDate3 = "#" + str3(1) + "/" + str3(2) + "/" + str3(0) + "#"  '实际结束日期
                        Dim dif1 = DateDiff(“d”, fromDate1, fromDate2)
                        Dim dif2 = DateDiff(“d”, fromDate2, fromDate3)
                        Dim dif3 = DateDiff(“d”, fromDate4, fromDate2)
                        If dif1 >= 0 And dif2 >= 0 Then
                            xlSheet.Cells(i, j).Interior.ColorIndex = 4  '设置单元格背景颜色
                            If dif3 > 0 Then
                                xlSheet.Cells(i, j).Interior.ColorIndex = 3  '设置单元格背景颜色
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        Dim xlSheet1 As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel页数
        Dim sheet_index = xlSheet.Index
        Globals.ThisAddIn.Application.Worksheets.Add(After:=xlSheet)

        xlSheet1 = Globals.ThisAddIn.Application.Worksheets(sheet_index + 1)
        xlSheet1.Cells(1, 1).value = "成员"
        xlSheet1.Cells(1, 2).value = "总任务"
        xlSheet1.Cells(1, 3).value = "已完成"
        xlSheet1.Cells(1, 4).value = "正在进行"

        '获取excel行数
        Dim n = getLineNumber(xlSheet, col, row)

        Dim a = 2
        Dim a_1 = 2
        For i = row To n + row - 2
            Dim line_n = get_line(xlSheet1, 2, 1)
            If line_n = 0 Then
                xlSheet1.Cells(a_1, 1).value = xlSheet.Cells(i, 4).value
                xlSheet1.Cells(a_1, 2).value = "1"
                If xlSheet.Cells(i, 3).value = "已实现" Then
                    xlSheet1.Cells(a_1, 3).value = "1"
                End If
                If xlSheet.Cells(i, 3).value = "实现中" Then
                    xlSheet1.Cells(a_1, 4).value = "1"
                End If
                a_1 += 1
            Else
                Dim b = 0
                For j = a To a + line_n
                    If xlSheet1.Cells(j, 1).value = xlSheet.Cells(i, 4).value Then
                        b += 1
                        Dim b_1 = Integer.Parse(xlSheet1.Cells(j, 2).value)
                        xlSheet1.Cells(j, 2).value = b_1 + 1
                        If xlSheet.Cells(i, 3).value = "已实现" Then
                            If CStr(xlSheet1.Cells(j, 3).value) = "" Then
                                xlSheet1.Cells(j, 3).value = "1"
                            Else
                                xlSheet1.Cells(j, 3).value = Integer.Parse(xlSheet1.Cells(j, 3).value) + 1
                            End If
                        End If
                        If xlSheet.Cells(i, 3).value = "实现中" Then
                            If CStr(xlSheet1.Cells(j, 4).value) = "" Then
                                xlSheet1.Cells(j, 4).value = "1"
                            Else
                                xlSheet1.Cells(j, 4).value = Integer.Parse(xlSheet1.Cells(j, 4).value) + 1
                            End If
                        End If
                    End If
                Next
                If b = 0 Then
                    xlSheet1.Cells(a_1, 1).value = xlSheet.Cells(i, 4).value
                    xlSheet1.Cells(a_1, 2).value = "1"
                    If xlSheet.Cells(i, 3).value = "已实现" Then
                        If CStr(xlSheet1.Cells(a_1, 3).value) = "" Then
                            xlSheet1.Cells(a_1, 3).value = "1"
                        Else
                            xlSheet1.Cells(a_1, 3).value = Integer.Parse(xlSheet1.Cells(a_1, 3).value) + 1
                        End If
                    End If
                    If xlSheet.Cells(i, 3).value = "实现中" Then
                        If CStr(xlSheet1.Cells(a_1, 4).value) = "" Then
                            xlSheet1.Cells(a_1, 4).value = "1"
                        Else
                            xlSheet1.Cells(a_1, 4).value = Integer.Parse(xlSheet1.Cells(a_1, 3).value) + 1
                        End If
                    End If
                    a_1 += 1
                End If
            End If
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim n = getLineNumber(xlSheet, col, row)

        '获取excel列数
        Dim col_index = 0
        For i = 1 To 10000
            If CStr(xlSheet.Cells(1, i).value) = "" Then
                col_index = i - 1
                Exit For
            End If
        Next

        Dim a = get_col_char(col_index - 1)

        Dim str1 = "A"
        Dim str2 = a
        '设置初始合并单元格
        Dim start_cell = 3
        '设置结束合并单元格
        Dim end_cell = 3
        For i = row To n + row - 1
            '合并单元格
            If i <> 3 Then
                If xlSheet.Cells(i, 1).value <> xlSheet.Cells(i - 1, 1).value Then
                    end_cell = i - 1
                    Dim str6 = str1.Trim() + Str(start_cell).Trim()
                    Dim str7 = str1.Trim() + Str(end_cell).Trim()
                    xlSheet.Range(str6.Trim() & ":".Trim() & str7.Trim()).MergeCells = True
                    start_cell = i
                End If
            End If
                '设置线条
                If xlSheet.Cells(i, 1).value <> xlSheet.Cells(i - 1, 1).value Then
                Dim str3 = str1.Trim() + Str(i).Trim()
                Dim str4 = str2.Trim() + Str(i).Trim()
                Dim str5 = str3.Trim() & ":".Trim() & str4.Trim()
                xlSheet.Range(str5).Borders(3).LineStyle = 1  '设置单元格顶部有线
                '1:左 2:右 3:顶 4:底 5:斜\ 6:斜/

            End If
        Next

    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        If xlSheet.Cells(1, 1).value = "一天" Then
            MsgBox("不能重复点击")
        Else
            '设置上方的颜色对照表格
            xlSheet.Cells(1, 1).value = "一天"
            xlSheet.Cells(1, 2).Interior.ColorIndex = color_1  '设置单元格背景颜色
            xlSheet.Cells(2, 1).value = "二天"
            xlSheet.Cells(2, 2).Interior.ColorIndex = color_2  '设置单元格背景颜色
            xlSheet.Cells(3, 1).value = "三天"
            xlSheet.Cells(3, 2).Interior.ColorIndex = color_3  '设置单元格背景颜色
            xlSheet.Cells(4, 1).value = "四天"
            xlSheet.Cells(4, 2).Interior.ColorIndex = color_4  '设置单元格背景颜色
            xlSheet.Cells(5, 1).value = "五天"
            xlSheet.Cells(5, 2).Interior.ColorIndex = color_5  '设置单元格背景颜色
            xlSheet.Cells(6, 1).value = "六天"
            xlSheet.Cells(6, 2).Interior.ColorIndex = color_6  '设置单元格背景颜色
            xlSheet.Cells(7, 1).value = "七天"
            xlSheet.Cells(7, 2).Interior.ColorIndex = color_7  '设置单元格背景颜色

            '合并单元格
            '部门单元格
            merge_cell(xlSheet, row_work, col_work, "部门")
            '成员单元格
            merge_cell(xlSheet, row_work, col_work + 1, "成员")
            '参与项目单元格
            merge_cell(xlSheet, row_work, col_work + 2, "参与项目")
            'bug数量单元格
            merge_cell(xlSheet, row_work, col_work + 3, "bug数量")
            '备注单元格
            merge_cell(xlSheet, row_work, col_work + 4, "备注")

            '设置当天之前7天和之后30天的时间
            xlSheet.Cells(row_work, col_work + 7).value = "=today()"
            Dim day1 = xlSheet.Cells(row_work, col_work + 7).value
            day1 = day1.AddDays(-7)
            For i = col_work + 5 To col_work + 79 Step 2
                Dim day2 As DateTime = day1.AddDays(1)
                xlSheet.Cells(row_work, i).value = day2
                day1 = day2
                xlSheet.Cells(row_work, i + 1).value = "=TEXT(""" + day2 + """," + """AAA""" + ")"
                xlSheet.Cells(row_work + 1, i).value = "总工时"
                xlSheet.Cells(row_work + 1, i + 1).value = "已完成"
            Next
            '自动筛选
            Dim str1 = get_col_char(col_work - 1)
            Dim str2 = str1.Trim() + Str(row_work + 1).Trim()
            Dim str3 = get_col_char(col_work + 79)
            Dim str4 = str3.Trim() + Str(row_work + 1).Trim()
            Dim str5 = str2.Trim() & ":".Trim() & str4.Trim()
            xlSheet.Range(str5).AutoFilter(Field:=1)
        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '固定把工时放在第二页
        Dim xlSheet1 As Excel.Worksheet
        xlSheet1 = Globals.ThisAddIn.Application.Worksheets(2)

        '获取excel行数
        Dim n = getLineNumber(xlSheet, col, row)

        Dim a = row_work + 2
        '代表成员那行
        Dim a_1 = col_work + 2
        For i = row To n + row - 2
            Dim line_n = get_line(xlSheet1, row_work + 2, col_work + 1)

            '第一次还没有记录
            If line_n = 0 Then
                'copy成员名
                xlSheet1.Cells(a_1, col_work + 1).value = xlSheet.Cells(i, 4).value
                'copy参与项目名
                If xlSheet1.Cells(a_1, col_work + 2).value = "" Then
                    xlSheet1.Cells(a_1, col_work + 2).value = xlSheet.Cells(i, p_col5).value
                ElseIf InStr(CStr(xlSheet.Cells(a_1, col_work + 2).Value), xlSheet.Cells(i, p_col5).value) Then
                    xlSheet1.Cells(a_1, col_work + 2).value = xlSheet1.Cells(a_1, col_work + 2).value
                Else
                    xlSheet1.Cells(a_1, col_work + 2).value = xlSheet1.Cells(a_1, col_work + 2).value + " + " + xlSheet.Cells(i, p_col5).value
                End If
                'copy工时

                '填充颜色
                If xlSheet.Cells(i, 3).value = "已实现" Then
                    xlSheet1.Cells(a_1, 3).value = "1"
                End If
                If xlSheet.Cells(i, 3).value = "实现中" Then
                    xlSheet1.Cells(a_1, 4).value = "1"
                End If
                a_1 += 1
            Else
                Dim b = 0
                For j = a To a + line_n
                    If xlSheet1.Cells(j, 1).value = xlSheet.Cells(i, 4).value Then
                        b += 1
                        Dim b_1 = Integer.Parse(xlSheet1.Cells(j, 2).value)
                        xlSheet1.Cells(j, 2).value = b_1 + 1
                        If xlSheet.Cells(i, 3).value = "已实现" Then
                            If CStr(xlSheet1.Cells(j, 3).value) = "" Then
                                xlSheet1.Cells(j, 3).value = "1"
                            Else
                                xlSheet1.Cells(j, 3).value = Integer.Parse(xlSheet1.Cells(j, 3).value) + 1
                            End If
                        End If
                        If xlSheet.Cells(i, 3).value = "实现中" Then
                            If CStr(xlSheet1.Cells(j, 4).value) = "" Then
                                xlSheet1.Cells(j, 4).value = "1"
                            Else
                                xlSheet1.Cells(j, 4).value = Integer.Parse(xlSheet1.Cells(j, 4).value) + 1
                            End If
                        End If
                    End If
                Next
                If b = 0 Then
                    xlSheet1.Cells(a_1, 1).value = xlSheet.Cells(i, 4).value
                    xlSheet1.Cells(a_1, 2).value = "1"
                    If xlSheet.Cells(i, 3).value = "已实现" Then
                        If CStr(xlSheet1.Cells(a_1, 3).value) = "" Then
                            xlSheet1.Cells(a_1, 3).value = "1"
                        Else
                            xlSheet1.Cells(a_1, 3).value = Integer.Parse(xlSheet1.Cells(a_1, 3).value) + 1
                        End If
                    End If
                    If xlSheet.Cells(i, 3).value = "实现中" Then
                        If CStr(xlSheet1.Cells(a_1, 4).value) = "" Then
                            xlSheet1.Cells(a_1, 4).value = "1"
                        Else
                            xlSheet1.Cells(a_1, 4).value = Integer.Parse(xlSheet1.Cells(a_1, 3).value) + 1
                        End If
                    End If
                    a_1 += 1
                End If
            End If
        Next
    End Sub
End Class
