Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Dim objExcel As Object
    Dim objWorkBooks As Object
    Dim objWorkSheet As Object
    Dim objExcel1 As Object
    Dim objWorkBooks1 As Object
    Dim objWorkSheet1 As Object
    Dim pathExcel As String
    Dim DstPath As String
    Dim DstExcelPath As String
    Dim ExcelName As String
    Dim SearchSheet As String
    Dim FindContext As String
    Dim srcWidth As Integer
    Dim srcHeight As Integer
    Dim SheetIndex As Integer
    Dim PageBuffer() As String '不定数组
    Dim VersionBuffer() As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim OpenFileDialog1 = New OpenFileDialog
        OpenFileDialog1.Title = "请选择文件"
        OpenFileDialog1.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            pathExcel = OpenFileDialog1.FileName
            ExcelName = OpenFileDialog1.SafeFileName
            'Dim nLength As Integer
            'nLength = pathExcel.Length - ExcelName.Length
            'OpenPath = pathExcel.Remove(nLength, ExcelName.Length)
            TextBox1.Text = pathExcel
        End If
        objExcel = CreateObject("Excel.Application")
        If Dir(pathExcel) <> "" Then
            objWorkBooks = objExcel.Workbooks.Open(pathExcel)
            objExcel.Visible = False '不显示excel
        Else
            MsgBox("不存在Excel!")
        End If
        
        SheetNum.Text = objExcel.Worksheets.Count

    End Sub

    Private Sub OpenSheet1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenSheet1.Click
        TextBox4.Clear()
        TextBox5.Clear()
        Dim Sheet_num As Integer
        Sheet_num = CInt(SheetNum.Text) 'string to integer

        If SheetIndex < 1 Or SheetIndex > Sheet_num Then
            OpenSheetResult1.Text = "失败"
        Else
            objWorkSheet = objWorkBooks.Sheets(SheetIndex)

            srcHeight = objWorkSheet.UsedRange.Rows.Count '获取Sheet表行数
            For i = 9 To srcHeight Step 1
                If objWorkSheet.cells(i, 1).value = "" Then
                    srcHeight = i
                    Exit For
                End If
            Next
            TextBox4.Text = srcHeight
            srcWidth = objWorkSheet.UsedRange.Columns.Count '获取Sheet表列数
            'For i = 0 To srcWidth Step 1
            '    If objWorkSheet.cells(9, i).value = "" Then
            '        srcWidth = i
            '        Exit For
            '    End If
            'Next
            TextBox5.Text = srcWidth
            OpenSheetResult1.Text = "成功"
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If objExcel Is Nothing Then

        Else
            objExcel.DisplayAlerts = False    '关闭时不提示保存
            objWorkSheet = Nothing
            objWorkBooks.close()
            objWorkBooks = Nothing
            objExcel.Quit()    '关闭EXCEL
        End If

    End Sub
    Dim SuccessNums As Integer
    Dim FailNums As Integer
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim f1 As New List(Of IO.FileInfo)
        Dim AName As String
        Dim OpenName As String
        AName = "TestResult"
        If DstPath Is Nothing Then
            MsgBox("请先输入目标路径!")
            GoTo AA
        End If
        Dim SuccessNums As Integer
        Dim FailNums As Integer
        For i = 0 To PageBuffer.Count - 1 Step 1
            DstExcelPath = DstPath + "\" + PageBuffer(i) + "\"
            If Dir((DstExcelPath)) <> "" Then
                Dim di As New IO.DirectoryInfo(DstExcelPath)
                f1 = di.GetFiles("*.xlsx", IO.SearchOption.AllDirectories).ToList

                For j = 0 To f1.Count - 1 Step 1
                    If InStr(f1(j).Name, AName) Then
                        OpenName = DstExcelPath + f1(j).ToString
                        If Dir((OpenName)) <> "" Then
                            OpenName = OpenName.Remove(OpenName.Length - 5, 5)
                            objExcel1 = CreateObject("Excel.Application")
                            objWorkBooks1 = objExcel1.Workbooks.Open(OpenName)
                            objExcel1.Visible = False

                            objWorkSheet1 = objWorkBooks1.Sheets(1) '打开第一页
                            objWorkSheet1.cells(3, 9).value = VersionBuffer(i)

                            objWorkBooks1.save()
                            ListBox1.Items.Add(PageBuffer(i) + " *.xlsx changed successful!")
                            objExcel1.DisplayAlerts = False    '关闭时不提示保存
                            objWorkSheet1 = Nothing
                            objWorkBooks1.close()
                            objWorkBooks1 = Nothing
                            objExcel1.Quit()    '关闭EXCEL
                            objExcel1 = Nothing
                            SuccessNums = SuccessNums + 1
                        Else
                            ListBox1.Items.Add(PageBuffer(i) + " *.xlsx can not find!")
                            FailNums = FailNums + 1
                        End If
                    End If
                Next
            Else
                ListBox1.Items.Add(PageBuffer(i) + "File can not find!")
            End If
            
        Next
        Dim msgbox1 As String
        msgbox1 = "替换完成 " + "成功数量：" + CStr(SuccessNums) + "  失败数量：" + CStr(FailNums)
        MsgBox(msgbox1)
AA:
    End Sub


    Private Sub Sheet_Index_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sheet_Index.TextChanged
        SheetIndex = Sheet_Index.Text
    End Sub

    Private Sub LoadDstPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadDstPath.Click
        DstPath = Save_DstPath.Text
    End Sub

    Private Sub Save_Page_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save_Page.Click
        If objWorkSheet.cells(8, 1).value = "页码" Then
            ReDim PageBuffer(0 To srcHeight - 9)
            For i = 0 To PageBuffer.Count - 1 Step 1
                PageBuffer(i) = objWorkSheet.cells(i + 9, 1).value
                'If PageBuffer(i).EndsWith(" ") Then
                '    PageBuffer(i) = PageBuffer(i).Remove(PageBuffer(i).Length - 1, 1)
                'End If
            Next
            ShowResult.Text = "页码缓存成功"
        Else
            MsgBox("无法找到页码!")
        End If

        If objWorkSheet.cells(8, 3).value = "版本" Then
            ReDim VersionBuffer(0 To srcHeight - 9)
            For i = 0 To VersionBuffer.Count - 1 Step 1
                VersionBuffer(i) = objWorkSheet.cells(i + 9, 3).value
            Next
            ShowResult.Text = "版本缓存成功"
        Else
            MsgBox("无法找到版本!")
        End If

    End Sub
    Dim nCount As Integer
    Private Sub ChangeOneByOne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeOneByOne.Click

        If nCount >= PageBuffer.Count Then
            Dim msgbox1 As String
            msgbox1 = "替换完成 " + "成功数量：" + CStr(SuccessNums) + "  失败数量：" + CStr(FailNums)
            MsgBox(msgbox1)
            GoTo AA
        End If
        PageText.Text = PageBuffer(nCount)
        VersionText.Text = VersionBuffer(nCount)
        DstExcelPath = DstPath + "\" + PageBuffer(nCount) + "\usecase"
        If Dir((DstExcelPath + ".xlsx")) <> "" Then
            objExcel1 = CreateObject("Excel.Application")
            objWorkBooks1 = objExcel1.Workbooks.Open(DstExcelPath)
            objExcel1.Visible = False

            objWorkSheet1 = objWorkBooks1.Sheets(1) '打开第一页
            objWorkSheet1.cells(3, 9).value = VersionBuffer(nCount)

            objWorkBooks1.save()
            ListBox1.Items.Add(PageBuffer(nCount) + " changed successful!")
            objExcel1.DisplayAlerts = False    '关闭时不提示保存
            objWorkSheet1 = Nothing
            objWorkBooks1.close()
            objWorkBooks1 = Nothing
            objExcel1.Quit()    '关闭EXCEL
            objExcel1 = Nothing
            SuccessNums = SuccessNums + 1
        Else
            ListBox1.Items.Add(PageBuffer(nCount) + " can not find!")
            FailNums = FailNums + 1
        End If
        nCount = nCount + 1
AA:
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        objExcel1 = CreateObject("Excel.Application")
        objWorkBooks1 = objExcel1.Workbooks.Open("F:\\demo.xlsx")
        objExcel1.Visible = False

        objWorkSheet1 = objWorkBooks1.Sheets(1) '打开第一页
        objWorkSheet1.select()
        objWorkBooks1.save()
        objExcel1.DisplayAlerts = False    '关闭时不提示保存
        objWorkSheet1 = Nothing
        objWorkBooks1.close()
        objWorkBooks1 = Nothing
        objExcel1.Quit()    '关闭EXCEL
        objExcel1 = Nothing
    End Sub
End Class
