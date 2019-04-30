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
    '设置路径
    Public path_address = "C:\Users\Public\Documents"
    '设置文件保存路径
    Public filename_path = "C:\Users\12959\测试资料\xmind\"
    '设置excel路径
    Public path = path_address + "\testcase.xlsx"
    '设置database地址
    Public data_address = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_address + "\testcase.accdb"
    Public Function write_title(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)   '设置首行title

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        xlSheet.Cells(1, col1).value = "用例目录"
        xlSheet.Cells(1, col2).value = "用例名称"
        xlSheet.Cells(1, col3).value = "需求ID"
        xlSheet.Cells(1, col4).value = "前置条件"
        xlSheet.Cells(1, col5).value = "用例步骤"
        xlSheet.Cells(1, col6).value = "预期结果"
        xlSheet.Cells(1, col7).value = "用例类型"
        xlSheet.Cells(1, col8).value = "用例状态"
        xlSheet.Cells(1, col9).value = "用例等级"
        xlSheet.Cells(1, col10).value = "创建人"

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
            ElseIf InStr(CStr(xlSheet.Cells(row1, i).Value), "。") Or InStr(CStr(xlSheet.Cells(row1, i).Value), ".") Then
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
    Public Function getLineNumber(xlsheet, num)  '获取excel行数函数

        Dim a = 2
        For k = 2 To 1000
            If CStr(xlsheet.Cells(k, num).value) = "" Then
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
        Else
            'MsgBox(arr(0))
            'MsgBox(arr(1))
            If InStr(arr(0), "H5") Or InStr(arr(0), "小程序") Then
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
                        ElseIf InStr(CStr(xlSheet.Cells(i, col_text).Value), ".") Then
                            arr = Split(xlSheet.Cells(i, col_text).value, ".")
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

            '清空copy进来的内容
            If a <> 0 Then

                clear_text(a, blank_front, row_num, b)

            End If


            '设置首行title
            write_title(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)


            '隐藏用例目录、需求ID
            xlSheet.Columns(col1).Hidden = True
            xlSheet.Columns(col3).Hidden = True

            xlSheet.Cells(1, 1).Interior.ColorIndex = 0  '设置单元格背景颜色
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
        Dim a = getLineNumber(xlSheet, col2)
        'MsgBox(a)
        '判断是否用例名称包含】
        Dim flag_title = True

        Dim sql_text(a) As String   '初始化sql语句
        For i = 2 To a + 1

            If InStr(CStr(xlSheet.Cells(i, col2).value), "】") Or InStr(CStr(xlSheet.Cells(i, col2).value), "]") Then
                If InStr(LCase(CStr(xlSheet.Cells(i, col2).value)), "pc") Then
                    sql_text(i - 2) = get_sqltext(CStr(xlSheet.Cells(i, col2).value), "PC")
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

        xlApp.Visible = False
        xlApp.Workbooks.Open(path)

        xlBook = xlApp.Workbooks(1)
        xlSheet1 = xlBook.Sheets("PC")
        xlSheet2 = xlBook.Sheets("H5-小程序")

        '获取excel行数
        Dim a1 = getLineNumber(xlSheet1, 2)
        Dim a2 = getLineNumber(xlSheet2, 2)

        Dim arr1(a1, 1) As String
        Dim arr2(a2, 1) As String

        If a1 <> 0 Then
            For i = 0 To a1 - 1
                arr1(i, 0) = "select step from PC where title = '" + Trim(xlSheet1.Cells(i + 2, 2).value) + "'"
                arr1(i, 1) = "insert into PC(title,step) values('" + xlSheet1.Cells(i + 2, 2).value + "','" + xlSheet1.Cells(i + 2, 3).value + "')"
            Next
            updateaccess(arr1, a1)
        End If

        If a2 <> 0 Then
            For i = 0 To a2 - 1
                arr2(i, 0) = "select step from H5 where title = '" + Trim(xlSheet2.Cells(i + 2, 2).value) + "'"
                arr2(i, 1) = "insert into H5(title,step) values('" + xlSheet2.Cells(i + 2, 2).value + "','" + xlSheet2.Cells(i + 2, 3).value + "')"
            Next
            updateaccess(arr2, a2)
        End If


        xlApp.Quit()
        GC.Collect()

    End Sub
End Class
