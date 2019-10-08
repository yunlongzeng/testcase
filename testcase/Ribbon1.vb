Imports Microsoft.Office.Tools.Ribbon
Imports System.Data.OleDb
Imports System.Data
Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient


Public Class Ribbon1

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
    '设置处理人所在的列数
    Public p_col8 = 4

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
    '数据库连接字符串
    Public connStr As String = "database=ta;server=106.13.175.210;Uid=root;Pwd=901207;charset=utf8;"
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
    Public Function fill_format(xlSheet, title， day1, n1, n2, stp)
        '设置上方的颜色对照表格
        xlSheet.Cells(1, 1).value = ”同一天任务数1"
        xlSheet.Cells(1, 2).Interior.ColorIndex = color_1  '设置单元格背景颜色
        xlSheet.Cells(2, 1).value = "同一天任务数2"
        xlSheet.Cells(2, 2).Interior.ColorIndex = color_2  '设置单元格背景颜色
        xlSheet.Cells(3, 1).value = "同一天任务数3"
        xlSheet.Cells(3, 2).Interior.ColorIndex = color_3  '设置单元格背景颜色
        xlSheet.Cells(4, 1).value = "同一天任务数4"
        xlSheet.Cells(4, 2).Interior.ColorIndex = color_4  '设置单元格背景颜色
        xlSheet.Cells(5, 1).value = "同一天任务数5"
        xlSheet.Cells(5, 2).Interior.ColorIndex = color_5  '设置单元格背景颜色
        xlSheet.Cells(6, 1).value = "同一天任务数6"
        xlSheet.Cells(6, 2).Interior.ColorIndex = color_6  '设置单元格背景颜色
        xlSheet.Cells(7, 1).value = "同一天任务数7"
        xlSheet.Cells(7, 2).Interior.ColorIndex = color_7  '设置单元格背景颜色
        xlSheet.Cells(1, 3).value = title

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

        '设置时间
        For i = n1 + 5 To n2 Step stp
            Dim day2 As DateTime = day1.AddDays(1)
            xlSheet.Cells(row_work, i).value = day2
            day1 = day2
            xlSheet.Cells(row_work, i + 1).value = "=TEXT(""" + day2 + """," + """AAA""" + ")"
            xlSheet.Cells(row_work + 1, i).value = "总工时"
            xlSheet.Cells(row_work + 1, i + 1).value = "已完成"
        Next
        '自动筛选
        Dim str1 = get_col_char(n1 - 1)
        Dim str2 = str1.Trim() + Str(row_work + 1).Trim()
        Dim str3 = get_col_char(n2)
        Dim str4 = str3.Trim() + Str(row_work + 1).Trim()
        Dim str5 = str2.Trim() & ":".Trim() & str4.Trim()
        xlSheet.Range(str5).AutoFilter(Field:=1)

    End Function
    Public Function copy_data(xlSheet, a_1, fromDate1, fromDate2, a, i, n1, n2, stp, str33, str44, str55, str66, str77, str88, str11)

        '第2页是否有记录
        Dim line_n = get_line(xlSheet, row_work + 2, col_work + 1)

        '第一次还没有记录
        If line_n = 0 Then
            'copy部门
            xlSheet.Cells(a_1, col_work).value = str88

            'copy成员名
            xlSheet.Cells(a_1, col_work + 1).value = str77

            'copy参与项目名
            xlSheet.Cells(a_1, col_work + 2).value = str11

            '与时间表对比后填充工时和颜色
            For j = n1 + 5 To n2 Step stp
                Dim str3() = Split(xlSheet.Cells(row_work, j).value, "/")
                'MsgBox(xlSheet.Cells(i, p_col7).value)
                Dim fromDate3 = "#" + str3(1) + "/" + str3(2) + "/" + str3(0) + "#"  '表格日期
                Dim dif3 = DateDiff("d", fromDate1, fromDate3)   '表格日期与开始日期差
                Dim dif4 = DateDiff("d", fromDate3, fromDate2)   '结束日期与表格日期差

                'copy工时到开始日期栏
                If dif3 = 0 Then
                    xlSheet.Cells(a_1, j).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                    xlSheet.Cells(a_1, j).value = "+" + str44
                    '如果状态是已完成，则工时填充至已完成列
                    If str33 = "已完成" Or str33 = "已实现" Then
                        xlSheet.Cells(a_1, j + 1).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                        xlSheet.Cells(a_1, j + 1).value = "-" + str44
                    End If
                End If

                '填充颜色
                If dif3 >= 0 And dif4 >= 0 Then
                    xlSheet.Cells(a_1, j).Interior.ColorIndex = color_1  '设置第一天的颜色
                    xlSheet.Cells(a_1, j + 1).Interior.ColorIndex = color_1  '设置第一天的颜色
                End If
            Next
            a_1 += 1

        Else     '有记录之后应该怎么操作
            Dim b = 0
            For j = a To a + line_n

                '判断是否有此人
                If InStr(CStr(xlSheet.Cells(j, col_work + 1).Value), str77) Then  '包含此人
                    b += 1

                    'copy参与项目名
                    If xlSheet.Cells(j, col_work + 2).value = "" Then
                        xlSheet.Cells(j, col_work + 2).value = str11
                    ElseIf InStr(CStr(xlSheet.Cells(j, col_work + 2).Value), str11) Then
                        Dim aa = "pass"
                    Else
                        xlSheet.Cells(j, col_work + 2).value = xlSheet.Cells(j, col_work + 2).value + " + " + str11
                    End If

                    '与时间表对比后填充工时和颜色
                    For k = n1 + 5 To n2 Step 2
                        Dim str3() = Split(xlSheet.Cells(row_work, k).value, "/")
                        Dim fromDate3 = "#" + str3(1) + "/" + str3(2) + "/" + str3(0) + "#"  '表格日期
                        Dim dif3 = DateDiff(“d”, fromDate1, fromDate3)   '表格日期与开始日期差
                        Dim dif4 = DateDiff(“d”, fromDate3, fromDate2)   '结束日期与表格日期差

                        'copy工时到开始日期栏
                        If dif3 = 0 Then
                            xlSheet.Cells(j, k).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                            xlSheet.Cells(j, k).value = CStr(xlSheet.Cells(j, k).value) + "+" + str44
                            'MsgBox(xlSheet1.Cells(j, k).value)
                            '如果状态是已完成，则工时填充至已完成列
                            If str33 = "已完成" Or str33 = "已实现" Then
                                xlSheet.Cells(j, k + 1).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                                xlSheet.Cells(j, k + 1).value += “-” + str44
                            End If
                        End If

                        '填充颜色,要判断颜色的code，一个一个叠加
                        If dif3 >= 0 And dif4 >= 0 Then
                            If xlSheet.Cells(j, k).Interior.ColorIndex = -4142 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_1  '设置第一天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_1  '设置第一天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_1 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_2  '设置第二天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_2  '设置第二天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_2 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_3  '设置第三天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_3  '设置第三天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_3 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_4  '设置第二天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_4  '设置第二天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_4 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_5  '设置第二天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_5  '设置第二天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_5 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_6  '设置第二天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_6  '设置第二天的颜色
                            ElseIf xlSheet.Cells(j, k).Interior.ColorIndex = color_6 Then
                                xlSheet.Cells(j, k).Interior.ColorIndex = color_7  '设置第二天的颜色
                                xlSheet.Cells(j, k + 1).Interior.ColorIndex = color_7  '设置第二天的颜色
                            Else
                                MsgBox("同一天任务超过7件！！！")
                                xlSheet.Cells(j, k).value = xlSheet.Cells(j, k).value + "+" + "同一天任务超过7件！！！"  '同一天任务超过7件报警
                            End If
                        End If

                    Next

                End If

            Next
            '不包含此人，就要新增
            If b = 0 Then
                'copy部门
                xlSheet.Cells(a + line_n, col_work).value = str88

                'copy成员名
                xlSheet.Cells(a + line_n, col_work + 1).value = str77

                'copy参与项目名
                xlSheet.Cells(a + line_n, col_work + 2).value = str11

                '与时间表对比后填充工时和颜色
                For j = n1 + 5 To n2 Step 2
                    Dim str3() = Split(xlSheet.Cells(row_work, j).value, "/")
                    Dim fromDate3 = "#" + str3(1) + "/" + str3(2) + "/" + str3(0) + "#"  '表格日期
                    Dim dif3 = DateDiff(“d”, fromDate1, fromDate3)   '表格日期与开始日期差
                    Dim dif4 = DateDiff(“d”, fromDate3, fromDate2)   '结束日期与表格日期差

                    'copy工时到开始日期栏
                    If dif3 = 0 Then
                        xlSheet.Cells(a + line_n, j).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                        xlSheet.Cells(a + line_n, j).value = “+” + str44
                        '如果状态是已完成，则工时填充至已完成列
                        If str33 = "已完成" Or str33 = "已实现" Then
                            xlSheet.Cells(a + line_n, j + 1).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
                            xlSheet.Cells(a + line_n, j + 1).value = “-” + str44
                        End If
                    End If

                    '填充颜色
                    If dif3 >= 0 And dif4 >= 0 Then
                        xlSheet.Cells(a + line_n, j).Interior.ColorIndex = color_1  '设置第一天的颜色
                        xlSheet.Cells(a + line_n, j + 1).Interior.ColorIndex = color_1  '设置第一天的颜色
                    End If
                Next

            End If
        End If
    End Function
    Public Function situ_iteration(xlSheet)
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
                If xlSheet.Cells(i, 1).value = "没有部门" Or xlSheet.Cells(i, 1).value = "产品" Then
                    Continue For
                End If
                xlSheet.Cells(i, p_col2).value = Replace(xlSheet.Cells(i, p_col2).value, "-", "/")
                xlSheet.Cells(i, p_col3).value = Replace(xlSheet.Cells(i, p_col3).value, "-", "/")
                Dim str() = Split(xlSheet.Cells(i, p_col2).value, "/")
                Dim str1() = Split(xlSheet.Cells(i, p_col3).value, "/")
                If (xlSheet.Cells(i, p_col4).value = "已完成" Or xlSheet.Cells(i, p_col4).value = "已实现") Then
                    xlSheet.Cells(i, p_col1).value = "已完成"
                    xlSheet.Cells(i, p_col1).Interior.ColorIndex = 4  '设置单元格背景颜色
                ElseIf str(0) = "" Or str1(0) = "" Then
                    xlSheet.Cells(i, p_col1).value = "未填写日期"
                    xlSheet.Cells(i, p_col1).Interior.ColorIndex = 15  '设置单元格背景颜色
                Else
                    Dim fromDate1 = "#" + str(1) + "/" + str(2) + "/" + str(0) + "#"
                    Dim dif1 = DateDiff(“d”, fromDate1, Now.Date)
                    Dim fromDate2 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"
                    Dim dif2 = DateDiff("d", fromDate2, Now.Date)

                    '任务完成，自动填充实际开始结束时间
                    If (xlSheet.Cells(i, p_col4).value = "已完成" Or xlSheet.Cells(i, p_col4).value = "已实现") Then
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
            '设置列宽
            xlSheet.Columns(2).ColumnWidth = 60
            '生成autofilter
            xlSheet.Range("A1:J1").AutoFilter(Field:=1)
            Threading.Thread.Sleep(500)   '等待
        End If

        Dim xl = Globals.ThisAddIn.Application
        xl.ActiveWindow.SplitColumn = 1
        xl.ActiveWindow.SplitRow = 0
        'xl.Application.ActiveWindow.FreezePanes = True
    End Function
    Public Function connectMysql(cmdText)

        '创建 SqlConnection 连接
        Dim conn As New MySqlConnection(connStr)
        conn.Open()

        Dim myCommand As MySqlCommand = New MySqlCommand(cmdText, conn)
        Dim strData(5000, 20) As String
        Dim i As Integer
        i = 0
        Dim ds As New DataSet
        Dim str = myCommand.ExecuteReader()
        Dim flag = True
        While flag
            If str.Read() Then
                For intindex As Integer = 0 To str.FieldCount - 1
                    strData(i, intindex) = str.Item(intindex).ToString
                Next
                i += 1
            Else
                flag = False
            End If
        End While
        conn.Close()
        Return strData
    End Function

    Public Function get_project_info(xlsheet, cmdText)
        If xlsheet.Cells(1, 1).value <> "" Then
            MsgBox("This sheet isn't blank!")
        Else
            xlsheet.Cells(1, 1).value = "部门"
            xlsheet.Cells(1, 2).value = "任务标题"
            xlsheet.Cells(1, 3).value = "状态"
            xlsheet.Cells(1, 4).value = "处理人"
            xlsheet.Cells(1, 5).value = "预计开始"
            xlsheet.Cells(1, 6).value = "预计结束"
            xlsheet.Cells(1, 7).value = "预估工时"

            Dim str = connectMysql(cmdText)

            xlsheet.Name = str(0, 11)
            If str(0, 0) = "" Then
                MsgBox("Iteration id is wrong!")
            Else
                xlsheet.Hyperlinks.Add(xlsheet.Cells(1, 2), str(0, 14))
                For i = 0 To 10000
                    If str(i, 0) = "" Then
                        Exit For
                    End If
                    For j = 0 To 20
                        If str(i, j) <> "" Then
                            If j = 8 Then  '第一列填充部门
                                xlsheet.Cells(i + 2, 1).value = str(i, j)
                                If str(i, j) = "没有部门" Or str(i, j) = "产品" Then
                                    Dim str11 = "A" + CStr(i + 2) + ":" + "I" + CStr(i + 2)
                                    xlsheet.Range(str11).Interior.ColorIndex = 24  '设置单元格背景颜色
                                End If
                            End If
                            If j = 2 Then  '第二列填充标题
                                xlsheet.Cells(i + 2, 2).value = str(i, j)
                            End If
                            If j = 3 Then  '第三列填充状态
                                xlsheet.Cells(i + 2, 3).value = str(i, j)
                            End If
                            If j = 7 Then  '第四列填充处理人
                                xlsheet.Cells(i + 2, 4).value = str(i, j)
                            End If
                            If j = 5 Then  '第五列填充预计开始
                                xlsheet.Cells(i + 2, 5).value = str(i, j)
                            End If
                            If j = 6 Then  '第六列填充预计结束
                                xlsheet.Cells(i + 2, 6).value = str(i, j)
                            End If
                            If j = 4 Then  '第七列填充预估工时
                                xlsheet.Cells(i + 2, 7).value = str(i, j)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                Next
                situ_iteration(xlsheet)

            End If

        End If

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
        situ_iteration(xlSheet)

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

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim cmdText = "select * from iteration where iterations_id='" & Trim(EditBox2.Text) & "'"
        If Trim(EditBox2.Text) <> "" Then
            get_project_info(xlSheet, cmdText)

        Else
            MsgBox("Please input iteration id!")
        End If

    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click

        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet

        If xlSheet.Cells(1, 1).Value <> "" Then
            MsgBox("This sheet isn't blank!")
            Exit Sub
        End If
        xlSheet.Name = "迭代总表"
        Dim cmdText = "select DISTINCT iterations_id,iteration_name,project_name,url from iteration where project_name in (select DISTINCT project_name from iteration)"
        Dim str = connectMysql(cmdText)
        xlSheet.Cells(1, 1).value = str(0, 2)
        xlSheet.Cells(1, 1).Interior.ColorIndex = 4  '设置单元格背景颜色

        Dim a = 1
        Dim b = 2
        For i = 0 To 1000
            If str(i, 0) = "" Then
                Exit For
            End If
            If str(i, 2) = xlSheet.Cells(1, a).value Then
                xlSheet.Cells(b, a).value = str(i, 1)
                xlSheet.Cells(b, a + 1).NumberFormatLocal = "@"
                xlSheet.Cells(b, a + 1).value = str(i, 0)
                xlSheet.Hyperlinks.Add(xlSheet.Cells(b, a), str(i, 3))
                b += 1
            Else
                a += 3
                xlSheet.Cells(1, a).value = str(i, 2)
                xlSheet.Cells(1, a).Interior.ColorIndex = 4  '设置单元格背景颜色
                b = 2
                i -= 1
            End If
        Next
        For i = 1 To 10
            xlSheet.Columns(i).AutoFit  '设置自适应列宽
        Next
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlSheet1 As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet

        For i = 2 To 100 Step 3
            If xlSheet.Cells(2, i).value = "" Then
                Exit For
            End If

            For j = 2 To 1000
                If xlSheet.Cells(j, i).value = "" Then
                    Exit For
                End If

                '获取excel页数
                Dim sheet_index = xlSheet.Index
                Globals.ThisAddIn.Application.Worksheets.Add(After:=xlSheet)
                xlSheet1 = Globals.ThisAddIn.Application.Worksheets(sheet_index + 1)
                Dim cmdText = "select * from iteration where iterations_id='" & xlSheet.Cells(j, i).value & "'"
                get_project_info(xlSheet1, cmdText)
                'Threading.Thread.Sleep(500)   '等待
            Next
        Next
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '**********************************combo box没有选择则直接退出********************************************
        If ComboBox1.Text = "" Then
            MsgBox("请选择部门！")
            Exit Sub
        ElseIf ComboBox1.Text <> "All" And ComboBox1.Text <> "测试" And ComboBox1.Text <> "前端" And ComboBox1.Text <> "UI" And ComboBox1.Text <> "产品" And ComboBox1.Text <> "全渠道" And ComboBox1.Text <> "CRM" And ComboBox1.Text <> "开放平台" And ComboBox1.Text <> "没有部门" Then
            MsgBox("输入的部门不正确！")
            Exit Sub
        End If

        '**********************************设置格式，统计一个月内的工时********************************************
        If CStr(xlSheet.Cells(1, 1).value) = "一天" Then
            MsgBox("不能重复点击")
            Exit Sub
        ElseIf CStr(xlSheet.Cells(1, 1).value) = "" Then

            '设置当天之前7天和之后30天的时间
            xlSheet.Cells(row_work, col_work + 7).value = "=today()"
            Dim day1 = xlSheet.Cells(row_work, col_work + 7).value
            day1 = day1.AddDays(-7)

            fill_format(xlSheet, "一个月内累计工时"， day1, col_work, col_work + 79, 2)

        Else
            MsgBox("此页面不为空，无法点击！")
            Exit Sub
        End If

        '**********************************excel进行填充********************************************
        Dim sql_text1 = ""
        If ComboBox1.Text = "All" Then
            sql_text1 = "select * from iteration"
        Else
            sql_text1 = "select * from iteration where dept = '" & ComboBox1.Text & "'"
        End If

        Dim str = connectMysql(sql_text1)    '获取数据库内容

        '代表成员那行
        Dim a = row_work + 2
        Dim a_1 = row_work + 2

        For i = 0 To 10000
            If str(i, 0) = "" Then
                Exit For
            End If
            '将--改成//
            str(i, 5) = Replace(str(i, 5), "-", "/")   '预计开始时间
            str(i, 6) = Replace(str(i, 6), "-", "/")   '预计结束时间
            '如果预估工时、预计开始、预计结束任一没有填写则跳过不计
            If str(i, 5) = "//" Or str(i, 6) = "//" Or str(i, 4) = "--" Then
                Continue For
            End If

            Dim str1() = Split(str(i, 5), "/")
            Dim fromDate1 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"  '开始日期
            Dim str2() = Split(str(i, 6), "/")
            Dim fromDate2 = "#" + str2(1) + "/" + str2(2) + "/" + str2(0) + "#"  '结束日期

            '判断当前日期与开始日期的差是否小于30，当前日期与结束日期差是否大于-7
            Dim dif1 = DateDiff(“d”, Now， fromDate1)   '当前日期与开始日期的差
            Dim dif2 = DateDiff(“d”, Now， fromDate2)   '当前日期与结束日期差
            If dif1 >= -7 And dif2 <= 30 Then

                copy_data(xlSheet, a_1, fromDate1, fromDate2, a, i, col_work, col_work + 79, 2, str(i, 3), str(i, 4), str(i, 5), str(i, 6), str(i, 7), str(i, 8), str(i, 11))

            End If
        Next

        '调整今日线
        '获取excel行数
        Dim n = getLineNumber(xlSheet, col_work + 1, row_work + 2)
        For i = col_work + 5 To col_work + 79 Step 2
            Dim str33() = Split(xlSheet.Cells(row_work, i).value, "/")
            'MsgBox(xlSheet.Cells(i, p_col7).value)
            Dim fromDate33 = "#" + str33(1) + "/" + str33(2) + "/" + str33(0) + "#"  '表格日期
            Dim dif11 = DateDiff(“d”, fromDate33, Now)
            If dif11 = 0 Then
                Dim a_text = get_col_char(i - 1)
                Dim str3 = a_text.Trim() + CStr(11)
                Dim str4 = a_text.Trim() + CStr(n + 1).Trim()
                Dim str5 = str3.Trim() & ":".Trim() & str4.Trim()
                xlSheet.Range(str5).Borders(2).LineStyle = 1  '设置单元格顶部有线
                '1:左 2:右 3:顶 4:底 5:斜\ 6:斜/
                Exit For
            End If

        Next

        '排版
        For i = col_work + 1 To col_work + 79 Step 2
            xlSheet.Columns(i).AutoFit  '设置自适应列宽
            '周六周日列宽度变窄
            If CStr(xlSheet.Cells(row_work, i + 1).value) = "六" Or CStr(xlSheet.Cells(row_work, i + 1).value) = "日" Then
                xlSheet.Columns(i).ColumnWidth = 1.5
                xlSheet.Columns(i + 1).ColumnWidth = 1.5
            End If
        Next
        xlSheet.Name = ComboBox1.Text
        MsgBox("Finished!")
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim cmdText = "select * from iteration where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bugcmdText = "select * from bug where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bug_num = "select count(*) from bug where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bug_fixed = "select count(*) from bug where iterations_id = '" & Trim(EditBox2.Text) & "' and (bug_status = '已解决' or bug_status = '已关闭')"
        If Trim(EditBox2.Text) <> "" Then
            '如果第一个单元格有文字则提示
            If xlSheet.Cells(1, 1).value <> "" Then
                MsgBox("This sheet isn't blank!")
                Exit Sub
            End If
            '读取数据库
            Dim str = connectMysql(cmdText)
            Dim str1 = connectMysql(bugcmdText)
            Dim bugnum = connectMysql(bug_num)
            Dim bugfixed = connectMysql(bug_fixed)
            '填充title
            If Now.Hour < 12 Then
                xlSheet.Cells(1, 1).value = Now.Date + "上午进度汇报"
            Else
                xlSheet.Cells(1, 1).value = Now.Date + "下午进度汇报"
            End If
            xlSheet.Cells(1, 2).value = str(0, 11)
            '填充本次发版任务
            xlSheet.Cells(2, 1).value = "本次发版任务："
            Dim n_tasks = 0   '总task
            Dim n_nostart = 0    '未开始的任务
            Dim n_sub = 0   '已提测的任务
            Dim n_fini = 0   '已测试结束的任务
            Dim a(50) As Integer
            a(0) = 0
            Dim j = 1

            Dim hui_date = ""
            For i = 0 To 10000
                If str(i, 0) = "" Then
                    Exit For
                End If
                If str(i, 8) = "没有部门" Or str(i, 8) = "产品" Then
                    n_tasks += 1
                    a(j) = i
                    j += 1
                End If
                If InStr(str(i, 2), "灰度") Then
                    a(j) = i
                    hui_date = str(i, 5)
                End If
            Next

            For i = 1 To a.Length
                If a(i + 1) = 0 Then
                    Exit For
                End If
                Dim flag1 = 1
                Dim flag2 = 1
                Dim flag3 = 1
                For j = a(i) + 1 To a(i + 1) - 1
                    If (str(j, 3) = "进行中" Or str(j, 3) = "已完成") And str(j, 8) <> "测试" Then
                        flag1 = 0
                        Exit For
                    End If
                Next
                For j = a(i) + 1 To a(i + 1) - 1
                    If (str(j, 3) = "未开始" Or str(j, 3) = "进行中") And str(j, 8) <> "测试" Then
                        flag2 = 0
                        Exit For
                    End If
                Next
                For j = a(i) + 1 To a(i + 1) - 1
                    If str(j, 3) = "未开始" Or str(j, 3) = "进行中" Then
                        flag3 = 0
                        Exit For
                    End If
                Next
                If flag1 = 1 Then
                    n_nostart += 1
                End If
                If flag2 = 1 Then
                    n_sub += 1
                End If
                If flag3 = 1 Then
                    n_fini += 1
                End If
            Next
            xlSheet.Cells(2, 2).value = CStr(n_tasks) + "个"
            '未开始
            xlSheet.Cells(3, 1).value = "未开始："
            xlSheet.Cells(3, 2).value = CStr(n_nostart) + "个"
            '开发中
            xlSheet.Cells(4, 1).value = "开发中："
            xlSheet.Cells(4, 2).value = CStr(n_tasks - n_nostart - n_sub) + "个"
            '已提测
            xlSheet.Cells(5, 1).value = "已提测："
            xlSheet.Cells(5, 2).value = CStr(n_sub) + "个"
            '测试通过
            xlSheet.Cells(6, 1).value = "测试通过："
            xlSheet.Cells(6, 2).value = CStr(n_fini) + "个"
            'bug总数
            xlSheet.Cells(7, 1).value = "bug总数："
            xlSheet.Cells(7, 2).value = bugnum(0, 0) + "个"
            '已修复：
            xlSheet.Cells(8, 1).value = "已修复："
            xlSheet.Cells(8, 2).value = bugfixed(0, 0) + "个"
            '未修复的阻碍bug：
            For i = 0 To 10000
                If str1(i, 0) = "" Then
                    Exit For
                End If
                If (str1(i, 2) = "新" Or str1(i, 2) = "重新打开" Or str1(i, 2) = "接受/处理") And (str1(i, 8) = "严重" Or str1(i, 8) = "致命") Then
                    If xlSheet.Cells(9, 2).value = "" Then
                        xlSheet.Cells(9, 2).value += str1(i, 1) + " @" + str1(i, 3)
                    Else
                        xlSheet.Cells(9, 2).value += Chr(10) + str1(i, 1) + " @" + str1(i, 3)
                    End If
                End If
            Next
            xlSheet.Cells(9, 1).value = "未修复的阻碍bug："
            '预计提测时间：
            xlSheet.Cells(10, 1).value = "预计提测时间："
            '预计上灰度时间：
            xlSheet.Cells(11, 1).value = "预计上灰度时间："
            xlSheet.Cells(11, 2).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
            xlSheet.Cells(11, 2).value = hui_date
            '预计发版时间：
            xlSheet.Cells(12, 1).value = "预计发版时间："
            xlSheet.Cells(12, 2).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
            xlSheet.Cells(12, 2).value = str(0, 13)
            '灰度测试情况
            xlSheet.Cells(13, 1).value = "灰度测试情况："
            '是否有延期：
            xlSheet.Cells(14, 1).value = "是否有延期？原因："
            '当天提测任务：
            xlSheet.Cells(15, 1).value = "当天提测任务："
            For i = 0 To 10000
                If str(i, 0) = "" Then
                    Exit For
                End If
                '将--改成//
                str(i, 6) = Replace(str(i, 6), "-", "/")   '预计结束时间
                '如果预估工时、预计开始、预计结束任一没有填写则跳过不计
                If str(i, 6) = "//" Then
                    Continue For
                End If
                Dim str2() = Split(str(i, 6), "/")
                Dim fromDate2 = "#" + str2(1) + "/" + str2(2) + "/" + str2(0) + "#"  '结束日期

                '判断延期任务
                Dim dif1 = DateDiff(“d”, Now.Date， fromDate2)   '当前日期与结束日期的差
                If dif1 = 0 And Not (str(i, 8) = "产品" Or str(i, 8) = "没有部门") And str(i, 8) <> "测试" Then
                    If xlSheet.Cells(15, 2).value = "" Then
                        xlSheet.Cells(15, 2).value += str(i, 2) + " @" + str(i, 7)
                    Else
                        xlSheet.Cells(15, 2).value += Chr(10) + str(i, 2) + " @" + str(i, 7)
                    End If
                End If
                If dif1 < 0 And Not (str(i, 8) = "产品" Or str(i, 8) = "没有部门" Or str(i, 8) = "测试") And str(i, 3) <> "已完成" Then
                    If xlSheet.Cells(16, 2).value = "" Then
                        xlSheet.Cells(16, 2).value += str(i, 2) + " @" + str(i, 7)
                    Else
                        xlSheet.Cells(16, 2).value += Chr(10) + str(i, 2) + " @" + str(i, 7)
                    End If

                End If
            Next
            '已延期的提测任务：
            xlSheet.Cells(16, 1).value = "已延期的提测任务："
            '需求变更：
            xlSheet.Cells(17, 1).value = "需求变更："
            '问题点：
            xlSheet.Cells(18, 1).value = "问题点："

            xlSheet.Columns(1).ColumnWidth = 20  '设置列宽
            xlSheet.Columns(2).ColumnWidth = 60  '设置列宽
            xlSheet.Rows(9).AutoFit  '设置自适应列宽
            xlSheet.Rows(15).AutoFit  '设置自适应列宽
            xlSheet.Rows(16).AutoFit  '设置自适应列宽
            xlSheet.Name = "日报"
            xlSheet.Cells(1, 1).Interior.ColorIndex = 33  '设置单元格背景颜色
            xlSheet.Cells(1, 2).Interior.ColorIndex = 33  '设置单元格背景颜色
            xlSheet.Cells(3, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(3, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(5, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(5, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(7, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(7, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(9, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(9, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(11, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(11, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(13, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(13, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(15, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(15, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(17, 1).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Cells(17, 2).Interior.ColorIndex = 15  '设置单元格背景颜色
            xlSheet.Range("A1:B18").Borders.LineStyle = 1
        Else
            MsgBox("Please input iteration id!")
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim cmdText = "select * from iteration where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bugcmdText = "select * from bug where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bug_num = "select count(*) from bug where iterations_id='" & Trim(EditBox2.Text) & "'"
        Dim bug_fixed = "select count(*) from bug where iterations_id = '" & Trim(EditBox2.Text) & "' and (bug_status = '已解决' or bug_status = '已关闭')"
        Dim hui_date = ""
        Dim n_tasks = 0
        Dim j = 1
        Dim a(50) As Integer
        Dim n_sub = 0
        Dim n_all = 0  '合计开发任务
        Dim n_subtask = 0  '已提测开发任务
        Dim n_openbug = 0  '未解决的缺陷

        If Trim(EditBox2.Text) <> "" Then
            '如果第一个单元格有文字则提示
            If xlSheet.Cells(1, 1).value <> "" Then
                MsgBox("This sheet isn't blank!")
                Exit Sub
            End If

            xlSheet.Name = "周报"
            '设置列宽
            xlSheet.Columns(2).ColumnWidth = 15
            xlSheet.Columns(3).ColumnWidth = 10
            xlSheet.Columns(4).ColumnWidth = 10
            xlSheet.Columns(5).ColumnWidth = 6
            xlSheet.Columns(6).ColumnWidth = 6
            xlSheet.Columns(7).ColumnWidth = 10
            xlSheet.Columns(8).ColumnWidth = 10
            xlSheet.Columns(9).ColumnWidth = 10
            xlSheet.Columns(10).ColumnWidth = 10
            xlSheet.Columns(11).ColumnWidth = 10
            xlSheet.Columns(12).ColumnWidth = 8
            xlSheet.Columns(13).ColumnWidth = 8
            xlSheet.Columns(14).ColumnWidth = 10
            xlSheet.Columns(15).ColumnWidth = 8
            xlSheet.Columns(16).ColumnWidth = 8
            '设置行高
            xlSheet.Rows(3).RowHeight = 15
            xlSheet.Rows(4).RowHeight = 15
            xlSheet.Rows(5).RowHeight = 15
            xlSheet.Rows(6).RowHeight = 15
            xlSheet.Rows(8).RowHeight = 30
            xlSheet.Rows(9).RowHeight = 15
            xlSheet.Rows(10).RowHeight = 15
            xlSheet.Rows(11).RowHeight = 15
            xlSheet.Rows(12).RowHeight = 15

            '读取数据库
            Dim str = connectMysql(cmdText)
            Dim str1 = connectMysql(bugcmdText)
            Dim bugnum = connectMysql(bug_num)
            Dim bugfixed = connectMysql(bug_fixed)
            '项目组成员
            xlSheet.Cells(3, 2).value = "项目组成员"
            xlSheet.Range("B3:B5").MergeCells = True
            xlSheet.Cells(3, 3).value = "产品经理："
            xlSheet.Cells(4, 3).value = "UI部门："
            xlSheet.Cells(5, 3).value = "前端："
            xlSheet.Cells(3, 8).value = "技术经理："
            xlSheet.Cells(4, 8).value = "测试："
            xlSheet.Cells(5, 8).value = "后端："
            xlSheet.Cells(3, 14).value = "项目经理："

            For i = 0 To 10000
                If str(i, 0) = "" Then
                    Exit For
                End If
                If InStr(str(i, 2), "灰度") Then
                    hui_date = str(i, 5)
                    a(j) = i
                End If
                If str(i, 8) = "产品" Then
                    If InStr(CStr(xlSheet.Cells(3, 4).value), str(i, 7)) Then
                        xlSheet.Cells(3, 4).value = xlSheet.Cells(3, 4).value
                    Else
                        xlSheet.Cells(3, 4).value += str(i, 7) + " "
                    End If
                End If
                If str(i, 8) = "UI" Then
                    If InStr(CStr(xlSheet.Cells(4, 4).value), str(i, 7)) Then
                        xlSheet.Cells(4, 4).value = xlSheet.Cells(4, 4).value
                    Else
                        xlSheet.Cells(4, 4).value += str(i, 7) + " "
                    End If
                    n_all += 1
                    If str(i, 3) = "已完成" Then
                        n_subtask += 1
                    End If
                End If
                    If str(i, 8) = "前端" Then
                    If InStr(CStr(xlSheet.Cells(5, 4).value), str(i, 7)) Then
                        xlSheet.Cells(5, 4).value = xlSheet.Cells(5, 4).value
                    Else
                        xlSheet.Cells(5, 4).value += str(i, 7) + " "
                    End If
                    n_all += 1
                    If str(i, 3) = "已完成" Then
                        n_subtask += 1
                    End If
                End If
                If str(i, 8) = "测试" Then
                    If InStr(CStr(xlSheet.Cells(4, 9).value), str(i, 7)) Then
                        xlSheet.Cells(4, 9).value = xlSheet.Cells(4, 9).value
                    Else
                        xlSheet.Cells(4, 9).value += str(i, 7) + " "
                    End If
                End If
                If str(i, 8) = "全渠道" Or str(i, 8) = "CRM" Or str(i, 8) = "开放平台" Then
                    If InStr(CStr(xlSheet.Cells(5, 9).value), str(i, 7)) Then
                        xlSheet.Cells(5, 9).value = xlSheet.Cells(5, 9).value
                    Else
                        xlSheet.Cells(5, 9).value += str(i, 7) + " "
                    End If
                    n_all += 1
                    If str(i, 3) = "已完成" Then
                        n_subtask += 1
                    End If
                End If
                If str(i, 8) = "没有部门" Or str(i, 8) = "产品" Then
                    n_tasks += 1
                    a(j) = i
                    j += 1
                End If

                '将--改成//
                str(i, 6) = Replace(str(i, 6), "-", "/")   '预计结束时间
                '如果预估工时、预计开始、预计结束任一没有填写则跳过不计
                If str(i, 6) = "//" Then
                    Continue For
                End If
                Dim str2() = Split(str(i, 6), "/")
                Dim fromDate2 = "#" + str2(1) + "/" + str2(2) + "/" + str2(0) + "#"  '结束日期

                '判断延期任务
                Dim dif1 = DateDiff(“d”, Now.Date， fromDate2)   '当前日期与结束日期的差
                If dif1 < 0 And Not (str(i, 8) = "产品" Or str(i, 8) = "没有部门" Or str(i, 8) = "测试") And str(i, 3) <> "已完成" Then
                    If xlSheet.Cells(11, 4).value = "" Then
                        xlSheet.Cells(11, 4).value += str(i, 2) + " @" + str(i, 7)
                    Else
                        xlSheet.Cells(11, 4).value += Chr(10) + str(i, 2) + " @" + str(i, 7)
                    End If
                    xlSheet.Rows(11).RowHeight += 10
                End If
            Next
            For i = 1 To a.Length
                If a(i + 1) = 0 Then
                    Exit For
                End If
                Dim flag2 = 1
                For j = a(i) + 1 To a(i + 1) - 1
                    If (str(j, 3) = "未开始" Or str(j, 3) = "进行中") And str(j, 8) <> "测试" Then
                        flag2 = 0
                        Exit For
                    End If
                Next
                If flag2 = 1 Then
                    n_sub += 1
                End If
            Next
            xlSheet.Range("D3:G3").MergeCells = True
            xlSheet.Range("D4:G4").MergeCells = True
            xlSheet.Range("D5:G5").MergeCells = True
            xlSheet.Range("I3:M3").MergeCells = True
            xlSheet.Range("O3:P3").MergeCells = True
            xlSheet.Range("I4:P4").MergeCells = True
            xlSheet.Range("I5:P5").MergeCells = True
            '风险
            xlSheet.Cells(9, 2).value = "风险"
            xlSheet.Range("B9:B12").MergeCells = True
            xlSheet.Cells(9, 3).value = "需求变更："
            xlSheet.Cells(10, 3).value = "严重bug："
            xlSheet.Cells(11, 3).value = "延期任务："
            xlSheet.Cells(12, 3).value = "上线延期："
            xlSheet.Range("D9:P9").MergeCells = True
            xlSheet.Range("D10:P10").MergeCells = True
            xlSheet.Range("D11:P11").MergeCells = True
            xlSheet.Range("D12:P12").MergeCells = True

            For i = 0 To 10000
                If str1(i, 0) = "" Then
                    Exit For
                End If
                If (str1(i, 2) = "新" Or str1(i, 2) = "重新打开" Or str1(i, 2) = "接受/处理") And (str1(i, 8) = "严重" Or str1(i, 8) = "致命") Then
                    If xlSheet.Cells(10, 4).value = "" Then
                        xlSheet.Cells(10, 4).value += str1(i, 1) + " @" + str1(i, 3)
                    Else
                        xlSheet.Cells(10, 4).value += Chr(10) + str1(i, 1) + " @" + str1(i, 3)
                    End If
                    xlSheet.Rows(10).RowHeight += 10
                End If
                If (str1(i, 2) = "新" Or str1(i, 2) = "重新打开" Or str1(i, 2) = "接受/处理") Then
                    n_openbug += 1
                End If
            Next
            '项目名称
            xlSheet.Cells(6, 2).value = "项目名称"
            xlSheet.Range("B6:B7").MergeCells = True
            xlSheet.Cells(8, 2).value = str(0, 11)
            xlSheet.Cells(8, 2).WrapText = True
            '灰度提测
            xlSheet.Cells(6, 3).value = "灰度提测"
            xlSheet.Range("C6:C7").MergeCells = True
            xlSheet.Cells(8, 3).value = hui_date
            '上线日期
            xlSheet.Cells(6, 4).value = "上线日期"
            xlSheet.Range("D6:D7").MergeCells = True
            xlSheet.Cells(8, 4).value = str(0, 13)
            '开发任务
            xlSheet.Cells(6, 5).value = "开发任务"
            xlSheet.Range("E6:F6").MergeCells = True
            xlSheet.Cells(7, 5).value = "已完成"
            xlSheet.Cells(8, 5).value = n_subtask
            xlSheet.Cells(7, 6).value = "合计"
            xlSheet.Cells(8, 6).value = n_all
            '提测情况
            xlSheet.Cells(6, 7).value = "提测情况"
            xlSheet.Range("G6:G7").MergeCells = True
            xlSheet.Cells(8, 7).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
            xlSheet.Cells(8, 7).value = CStr(n_sub) + "/" + CStr(n_tasks)
            '测试计划
            xlSheet.Cells(6, 8).value = "测试计划"
            xlSheet.Range("H6:K6").MergeCells = True
            xlSheet.Cells(7, 8).value = "用例设计"
            xlSheet.Cells(7, 9).value = "用例评审"
            xlSheet.Cells(7, 10).value = "测试环境测试"
            xlSheet.Cells(7, 10).WrapText = True
            xlSheet.Cells(7, 11).value = "灰度测试"
            '用例执行
            xlSheet.Cells(6, 12).value = "设计用例"
            xlSheet.Range("L6:L7").MergeCells = True
            xlSheet.Cells(6, 13).value = "执行用例"
            xlSheet.Range("M6:M7").MergeCells = True
            xlSheet.Cells(6, 14).value = "执行进度"
            xlSheet.Range("N6:N7").MergeCells = True
            xlSheet.Cells(8, 14).value = "=ROUND(M8/L8*100,0)&""%"""
            '缺陷
            xlSheet.Cells(6, 15).value = "缺陷总数"
            xlSheet.Range("O6:O7").MergeCells = True
            xlSheet.Cells(8, 15).value = bugnum(0, 0)
            xlSheet.Cells(6, 16).value = "未解决数"
            xlSheet.Range("P6:P7").MergeCells = True
            xlSheet.Cells(8, 16).value = n_openbug
            xlSheet.Range("B6:P12").Borders.LineStyle = 1
            xlSheet.Range("B3:B5").Borders.LineStyle = 1
            xlSheet.Range("C3:P3").Borders(3).LineStyle = 1  '设置单元格顶部有线
            xlSheet.Range("P3:P5").Borders(2).LineStyle = 1  '设置单元格顶部有线
            '1:左 2:右 3:顶 4:底 5:斜\ 6:斜/
            '加粗
            xlSheet.Cells(3, 2).Font.Bold = True
            xlSheet.Cells(8, 2).Font.Bold = True
            xlSheet.Cells(9, 2).Font.Bold = True
            xlSheet.Cells(9, 3).Font.Bold = True
            xlSheet.Cells(10, 3).Font.Bold = True
            xlSheet.Cells(11, 3).Font.Bold = True
            xlSheet.Cells(12, 3).Font.Bold = True
            xlSheet.Cells(9, 3).Font.ColorIndex = 3
            xlSheet.Cells(10, 3).Font.ColorIndex = 3
            xlSheet.Cells(11, 3).Font.ColorIndex = 3
            xlSheet.Cells(12, 3).Font.ColorIndex = 3
            xlSheet.Range("B6:P7").Interior.ColorIndex = 35  '设置单元格背景颜色
            xlSheet.Range("B6:P8").HorizontalAlignment = 3
            xlSheet.Range("B3:B11").HorizontalAlignment = 3

        Else
            MsgBox("Please input iteration id!")
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim cmdText = "select * from iteration where iterations_id='" & Trim(EditBox2.Text) & "'"
        If Trim(EditBox2.Text) <> "" Then
            '如果第一个单元格有文字则提示
            If xlSheet.Cells(1, 1).value <> "" Then
                MsgBox("This sheet isn't blank!")
                Exit Sub
            End If
            '读取数据库
            Dim str = connectMysql(cmdText)

            '*******************************
            '设置列宽
            xlSheet.Columns(1).ColumnWidth = 20
            xlSheet.Columns(2).ColumnWidth = 20
            xlSheet.Columns(3).ColumnWidth = 20
            xlSheet.Columns(4).ColumnWidth = 20
            '设置行高
            xlSheet.Rows(1).RowHeight = 30

            Dim hui_date = ""
            Dim row_fill = 6
            For i = 0 To 10000

                If str(i, 0) = "" Then
                    Exit For
                End If
                If InStr(str(i, 2), "灰度") Then
                    hui_date = str(i, 5)
                End If
                If str(i, 8) = "产品" Or str(i, 8) = "没有部门" Then
                    row_fill += 1
                    xlSheet.Cells(row_fill, 1) = str(i, 2)
                    Dim str_excel = "A" & CStr(row_fill) & ":D" & CStr(row_fill)
                    xlSheet.Range(str_excel).Interior.ColorIndex = 34  '设置单元格背景颜色
                    xlSheet.Range(str_excel).MergeCells = True
                    xlSheet.Cells(row_fill, 1).WrapText = True

                End If
                If str(i, 8) = "前端" Or str(i, 8) = "全渠道" Or str(i, 8) = "CRM" Or str(i, 8) = "开放平台" Then
                    row_fill += 1
                    xlSheet.Cells(row_fill, 1) = str(i, 7)
                    xlSheet.Cells(row_fill, 2) = "□Ucloud  □TP"
                End If

            Next
            '*************************************88
            '表抬头
            xlSheet.Name = "发灰度跟踪表"
            xlSheet.Cells(1, 1).value = "发灰度跟踪表"
            xlSheet.Range("A1:D1").MergeCells = True
            xlSheet.Cells(2, 1).value = "项目："
            xlSheet.Cells(2, 2).value = str(0, 11)
            xlSheet.Range("B2:D2").MergeCells = True
            xlSheet.Cells(3, 1).value = "灰度发布日期："
            xlSheet.Cells(3, 2).NumberFormatLocal = "@"  'G/通用格式,@为文本格式
            xlSheet.Cells(3, 2).value = hui_date
            xlSheet.Range("B3:D3").MergeCells = True
            xlSheet.Cells(6, 1).value = "开发人员"
            xlSheet.Cells(6, 2).value = "灰度环境"
            xlSheet.Cells(6, 3).value = "自测签名"
            xlSheet.Cells(6, 4).value = "日期"

            xlSheet.Cells(1, 1).Font.Bold = True
            xlSheet.Cells(2, 1).Font.Bold = True
            xlSheet.Cells(3, 1).Font.Bold = True
            xlSheet.Cells(6, 1).Font.Bold = True
            xlSheet.Cells(6, 2).Font.Bold = True
            xlSheet.Cells(6, 3).Font.Bold = True
            xlSheet.Cells(6, 4).Font.Bold = True

            xlSheet.Range("A1:D1").HorizontalAlignment = 3  '文字居中展示
            xlSheet.Range("A6:D6").HorizontalAlignment = 3

            Dim str_line = "A6:" & "D" & CStr(row_fill)
            xlSheet.Range(str_line).Borders.LineStyle = 1
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '**********************************Excel第一个单元格有值则直接退出********************************************
        If xlSheet.Cells(1, 1).value <> "" Then
            MsgBox("This sheet isn't blank！")
            Exit Sub
        ElseIf ComboBox2.Text <> "" And ComboBox2.Text <> "待开发" And ComboBox2.Text <> "开发中" And ComboBox2.Text <> "已实现" And ComboBox2.Text <> "挂起" And ComboBox2.Text <> "拒绝" Then
            MsgBox("Status is wrong！")
            Exit Sub
        ElseIf ComboBox3.Text <> "" And ComboBox3.Text <> "分销" And ComboBox3.Text <> "商城" And ComboBox3.Text <> "微信" Then
            MsgBox("Category is wrong！")
            Exit Sub
        End If

        '**********************************设置格式********************************************
        xlSheet.Cells(1, 1).value = "标题"
        xlSheet.Cells(1, 2).value = "状态"
        xlSheet.Cells(1, 3).value = "创建时间"
        xlSheet.Cells(1, 4).value = "是否超三周"
        xlSheet.Cells(1, 5).value = "期望上线"
        xlSheet.Cells(1, 6).value = "客户"
        xlSheet.Cells(1, 7).value = "创建人"
        xlSheet.Cells(1, 8).value = "处理人"
        xlSheet.Cells(1, 9).value = "优先级"
        xlSheet.Cells(1, 10).value = "迭代"
        xlSheet.Cells(1, 11).value = "细分类"
        xlSheet.Cells(1, 12).value = "分类"
        xlSheet.Cells(1, 13).value = "完成时间"

        '**********************************查找数据库********************************************
        Dim sql_text1 = ""
        If EditBox4.Text <> "" Then     '有结束时间则忽略其余参数，只考虑category
            If ComboBox3.Text = "" Then
                sql_text1 = "SELECT * FROM `stories` where finish_time LIKE '" + EditBox4.Text + "%'"
            Else
                sql_text1 = "SELECT * FROM `stories` where finish_time LIKE '" + EditBox4.Text + "%' and organization = '" + ComboBox3.Text + "'"
            End If
        Else  '没有结束时间
            If ComboBox2.Text = "" Then   'status为空的几种情况
                sql_text1 = "SELECT * FROM `stories`"
                If EditBox3.Text <> "" Then
                    sql_text1 += " where create_time LIKE '" + EditBox3.Text + "%'"
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                Else
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " where organization='" + ComboBox3.Text + "'"
                    End If
                End If
            Else   'status不为空的情况
                If ComboBox2.Text = "拒绝" Then
                    sql_text1 = "SELECT * FROM `stories` where status = '已拒绝'"
                    If EditBox3.Text <> "" Then
                        sql_text1 += " and create_time LIKE '" + EditBox3.Text + "%'"
                    End If
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                ElseIf ComboBox2.Text = "挂起" Then
                    sql_text1 = "SELECT * FROM `stories` where status = '挂起&延期'"
                    If EditBox3.Text <> "" Then
                        sql_text1 += " and create_time LIKE '" + EditBox3.Text + "%'"
                    End If
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                ElseIf ComboBox2.Text = "已实现" Then
                    sql_text1 = "SELECT * FROM `stories` where status = '已实现'"
                    If EditBox3.Text <> "" Then
                        sql_text1 += " and create_time LIKE '" + EditBox3.Text + "%'"
                    End If
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                ElseIf ComboBox2.Text = "开发中" Then
                    sql_text1 = "SELECT * FROM `stories` where status in ('原型设计','原型评审','需求评估','UI设计','UI评审','待开发','开发中','测试中') and iteration_title <>''"
                    If EditBox3.Text <> "" Then
                        sql_text1 += " and create_time LIKE '" + EditBox3.Text + "%'"
                    End If
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                ElseIf ComboBox2.Text = "待开发" Then
                    sql_text1 = "SELECT * FROM `stories` where status in ('原型设计','原型评审','需求评估','UI设计','UI评审','待开发','开发中','测试中') and iteration_title =''"
                    If EditBox3.Text <> "" Then
                        sql_text1 += " and create_time LIKE '" + EditBox3.Text + "%'"
                    End If
                    If ComboBox3.Text <> "" Then
                        sql_text1 += " and organization='" + ComboBox3.Text + "'"
                    End If
                Else
                    MsgBox("有异常情况！！！")
                End If
            End If

        End If

        Dim str = connectMysql(sql_text1)    '获取数据库内容

        '**********************************填充excel********************************************
        For i = 0 To 10000
            If Str(i, 0) = "" Then
                Exit For
            End If
            '标题+超级链接
            xlSheet.Cells(i + 2, 1).value = str(i, 2)
            xlSheet.Hyperlinks.Add(xlSheet.Cells(i + 2, 1), str(i, 13))
            '状态
            xlSheet.Cells(i + 2, 2).value = str(i, 7)
            '创建时间
            xlSheet.Cells(i + 2, 3).value = str(i, 9)
            '是否超三周

            str(i, 9) = Replace(str(i, 9), "-", "/")
            Dim str1() = Split(str(i, 9), "/")
            Dim fromDate1 = "#" + str1(1) + "/" + str1(2) + "/" + str1(0) + "#"
            Dim dif1 = DateDiff(“d”, fromDate1, Now.Date)
            If dif1 > 21 Then
                If str(i, 7) <> "已拒绝" And str(i, 7) <> "挂起&延期" And str(i, 7) <> "已实现" Then
                    If str(i, 5) = "" Then
                        xlSheet.Cells(i + 2, 4).value = "是"
                        xlSheet.Cells(i + 2, 4).Interior.ColorIndex = 3  '设置单元格背景颜色
                    Else
                        xlSheet.Cells(i + 2, 4).value = "迭代中"
                        xlSheet.Cells(i + 2, 4).Interior.ColorIndex = 7  '设置单元格背景颜色
                    End If

                End If
            End If


            '期望上线
            xlSheet.Cells(i + 2, 5).value = str(i, 12)
            '客户
            xlSheet.Cells(i + 2, 6).value = str(i, 11)
            '创建人
            xlSheet.Cells(i + 2, 7).value = str(i, 10)
            '处理人
            xlSheet.Cells(i + 2, 8).value = str(i, 8)
            '优先级
            xlSheet.Cells(i + 2, 9).value = str(i, 3)
            '迭代+超级链接
            xlSheet.Cells(i + 2, 10).value = str(i, 5)
            xlSheet.Hyperlinks.Add(xlSheet.Cells(i + 2, 10), str(i, 6))
            '细分类
            xlSheet.Cells(i + 2, 11).value = str(i, 4)
            '分类
            xlSheet.Cells(i + 2, 12).value = str(i, 14)
            '完成时间
            xlSheet.Cells(i + 2, 13).value = str(i, 15)
        Next
        For i = 1 To 13
            xlSheet.Columns(i).AutoFit  '设置自适应列宽
        Next

        '**********************************excel进行统计********************************************

        xlSheet.Range("P1:R1").MergeCells = True
        xlSheet.Cells(1, 16).Interior.ColorIndex = 35  '设置单元格背景颜色
        xlSheet.Cells(1, 16).value = "延期需求统计"
        xlSheet.Cells(2, 16).value = "未规划"
        xlSheet.Cells(2, 17).value = "已规划"
        xlSheet.Cells(2, 18).value = "分类"
        xlSheet.Cells(3, 16).value = "=COUNTIFS(D2:D10000,""是"",L2:L10000,R3)"
        xlSheet.Cells(3, 17).value = "=COUNTIFS(D2:D10000,""迭代中"",L2:L10000,R3)"
        xlSheet.Cells(3, 18).value = "商城"
        xlSheet.Cells(4, 16).value = "=COUNTIFS(D2:D10000,""是"",L2:L10000,R4)"
        xlSheet.Cells(4, 17).value = "=COUNTIFS(D2:D10000,""迭代中"",L2:L10000,R4)"
        xlSheet.Cells(4, 18).value = "分销"
        xlSheet.Cells(5, 16).value = "=COUNTIFS(D2:D10000,""是"",L2:L10000,R5)"
        xlSheet.Cells(5, 17).value = "=COUNTIFS(D2:D10000,""迭代中"",L2:L10000,R5)"
        xlSheet.Cells(5, 18).value = "微信"

        '设置列宽
        xlSheet.Columns(1).ColumnWidth = 60
        xlSheet.Columns(6).ColumnWidth = 10
        xlSheet.Columns(10).ColumnWidth = 40
        xlSheet.Range("A1:M1").AutoFilter(Field:=1)
    End Sub
End Class
