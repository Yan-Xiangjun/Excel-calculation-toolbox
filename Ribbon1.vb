'Copyright 2023 阎相君
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
    '开
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        For Each c In xls.Selection.Cells
            If c.Column >= 3 Then c.Interior.Color = 15132390
        Next
    End Sub
    '关
    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        xls.Selection.Interior.Pattern = -4142
    End Sub

    '单元格命名
    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim r_boundary = xls.Selection.Rows.Count
        Dim c_boundary = xls.Selection.Columns.Count
        Dim var, num, matches, unit
        Dim i As Integer
        Dim i_err As Integer
        Dim re = CreateObject("VBScript.RegExp") : re.Pattern = "[a-zA-Z]+(?=[0-9])"


        Dim cornerRD = xls.Selection.Cells(r_boundary, c_boundary).Value
        If r_boundary = 1 Then
            var = xls.Selection.Columns(1)
            If Left(cornerRD, 1) Like "[~a-zA-Z]" Then                 'a  1 (mm)
                num = xls.Selection.Columns(c_boundary - 1)
                unit = xls.Selection.Columns(c_boundary)
            Else                                              'a (1)
                num = xls.Selection.Columns(c_boundary)
            End If
        ElseIf c_boundary = 1 Then
            var = xls.Selection.Rows(1)
            If Left(cornerRD, 1) Like "[~a-zA-Z]" Then                 'a
                num = xls.Selection.Rows(r_boundary - 1)      '1
                unit = xls.Selection.Rows(r_boundary)         '(mm)
            Else
                num = xls.Selection.Rows(r_boundary)          'a
            End If                                            '(1)
        Else
            Dim cornerRD_left = xls.Selection.Cells(r_boundary, c_boundary - 1).Value
            If Left(cornerRD, 1) Like "[~a-zA-Z]" Then '有单位
                If VarType(cornerRD_left) = vbString Then
                    var = xls.Selection.Rows(1)                     'a   b   c
                    num = xls.Selection.Rows(r_boundary - 1)        '1   2   3
                    unit = xls.Selection.Rows(r_boundary)           'mm [N] (MPa)
                Else
                    var = xls.Selection.Columns(1)                  'a  1  mm
                    num = xls.Selection.Columns(c_boundary - 1)     'b  2  N
                    unit = xls.Selection.Columns(c_boundary)        'c [3] (MPa)
                End If
            Else '没有单位
                If VarType(cornerRD_left) = vbString Then
                    var = xls.Selection.Columns(1)                  'a    1
                    num = xls.Selection.Columns(c_boundary)         '[b] (2)
                Else
                    var = xls.Selection.Rows(1)                     'a   b   c
                    num = xls.Selection.Rows(r_boundary)            '1  [2] (3)
                End If
            End If
        End If

        For i = 1 To var.Cells.Count
            Dim n = var.Cells(i).Value
            matches = re.Execute(n)
            If matches.Count <> 0 Then
                var.Cells(i).Value = Left(n, matches(0).Length) & "_" & Mid(n, matches(0).Length + 1)
                n = var.Cells(i).Value
            End If

            Dim a = num.Cells(i).Address
            Try
                xls.ActiveWorkbook.Names.Add(n, "=" & a)
            Catch ex As Exception
                MsgBox(n & "所在单元格命名失败！" & vbCr & ex.Message, vbExclamation)
                i_err = i_err + 1
            End Try

            If unit IsNot Nothing Then
                Dim n_unit = n & "_unit"
                Dim a_unit = unit.Cells(i).Address
                xls.ActiveWorkbook.Names.Add(n_unit, "=" & a_unit)
            End If
        Next
        MsgBox("处理了" & i - 1 & "个变量，其中" & i - 1 - i_err & "个命名成功" & i_err & "个命名失败", vbInformation)
    End Sub

    '发送到Word
    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim doc = CreateObject("Word.Application")
        doc.Visible = True
        Dim wd = doc.documents.Add
        Dim s As String
        Dim ct As Integer
        For Each r In xls.Selection.Rows
            ct = r.Cells.Count
            If ct > 23 Then
                If MsgBox("选区内当前行有" & ct & "个单元格需要发送，耗时较长，确定要发送吗？", vbYesNo + vbQuestion + vbDefaultButton2) = 7 Then
                    doc.Quit
                    Exit Sub
                End If
            End If

            If has_smart_cell(r) Then
                Dim idx = unit_cell_index(r)
                If idx <= r.Cells.Count Then '该行有名字以“_unit”结尾的单元格

                    For i = idx To r.Cells.Count '单位及其后面的单元格
                        s = r.Cells(i).Text
                        If i = idx And s = "~" Then Continue For
                        If s <> "" Then doc.Selection.TypeText(s)
                    Next
                    Call formu_edit(doc, wd, "")

                    If idx >= 2 Then '单位前面的单元格（如果有的话）（一般有）
                        doc.Selection.TypeText(" ")
                        doc.Selection.MoveLeft(wdCharacter, 1)

                        For i = 1 To idx - 1
                            s = r.Cells(i).Text
                            If s <> "" Then doc.Selection.TypeText(s)
                        Next
                        Call formu_edit(doc, wd, "pro")

                    End If
                    Else '该行没有名字以“_unit”结尾的单元格
                    For Each c In r.Cells
                        s = c.Text
                        If s <> "" Then doc.Selection.TypeText(s)
                    Next
                    Call formu_edit(doc, wd, "pro")
                End If
                doc.Selection.MoveDown(wdParagraph, 1)
                doc.Selection.TypeText(vbCr)

            Else
                For Each c In r.Cells
                    s = c.Text
                    If s <> "" Then doc.Selection.TypeText(s)
                Next
                doc.Selection.TypeText(vbCr)
            End If

        Next
        MsgBox("完成！", vbInformation)
    End Sub
    Function has_smart_cell(rng As Excel.Range)
        For Each c In rng.Cells
            If c.Interior.Color = 15132390 Then
                has_smart_cell = True
                Exit Function
            End If
        Next
        has_smart_cell = False
    End Function
    Function unit_cell_index(rng As Excel.Range)
        Dim i As Integer
        For i = 1 To rng.Cells.Count
            Try
                If Right(rng.Cells(i).Name.Name, 5) = "_unit" Then
                    unit_cell_index = i
                    Exit Function
                End If
            Catch

            End Try
        Next
        unit_cell_index = i
    End Function
    Sub formu_edit(ByRef doc, ByRef wd, type)
        Dim ph_st = doc.Selection.Paragraphs(1).Range.Start
        Dim st = doc.Selection.Range.Start
        wd.Range(ph_st, st).Select

        Dim wd_rng = doc.Selection.Range
        If wd_rng.Text <> "" Then
            doc.Selection.OMaths.Add(wd_rng) '括号里不能直接填doc.Selection.Range，在后期绑定的情况下会报错
            doc.Selection.OMaths(1).Type = wdOMathInline
            doc.Selection.Font.Italic = False
            If type = "pro" Then doc.Selection.OMaths.BuildUp()
            doc.Selection.MoveLeft(wdCharacter, 2)
        End If
    End Sub

    '打开建标库
    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Shell("explorer.exe http://www.jianbiaoku.com/", vbNormalFocus)
    End Sub
    '刷新
    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim temp = Globals.Ribbons.Ribbon1.Chk5.Checked
        Dim choose = MsgBox("刷新单位？", vbQuestion + vbYesNoCancel + vbDefaultButton2)
        If choose = 6 Then
            Globals.Ribbons.Ribbon1.Chk5.Checked = True
        ElseIf choose = 7 Then
            Globals.Ribbons.Ribbon1.Chk5.Checked = False
        Else
            Exit Sub
        End If

        For Each c In xls.Selection.Cells
            c.Formula = c.Formula
        Next
        Globals.Ribbons.Ribbon1.Chk5.Checked = temp
    End Sub

    '打开项目地址
    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Shell("explorer.exe https://gitee.com/yan-xiangjun/excel-calculation-toolbox", vbNormalFocus)
    End Sub
    '可识别的公式
    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        Dim f1 As New excel_functions
        f1.Show()
        f1.TopMost = True
    End Sub
End Class
