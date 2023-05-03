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
Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Globals.Ribbons.Ribbon1.DropDown1.SelectedItemIndex = 2
        superscripts = {"²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹"}
        ws = CreateObject("WScript.Shell")
        py_location = Environment.GetEnvironmentVariable("Excel_calculation_toolbox", 2)
        my_function_xlam = xls.Workbooks.Open(py_location & "my_function.xlam")

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_SheetChange(Sh As Object, Target As Range) Handles Application.SheetChange
        xls.EnableEvents = False

        If Target.Cells.Count = 1 Then
            If Target.Interior.Color = 15132390 And Target.Formula <> "" Then
                xls.ScreenUpdating = False
                'ROUND
                Dim cell_formu_now = Target.Formula
                Dim d = Globals.Ribbons.Ribbon1.DropDown1.SelectedItem.Label
                If Globals.Ribbons.Ribbon1.Chk4.Checked Then
                    If Left(cell_formu_now, 1) = "=" Then
                        If Left(cell_formu_now, 7) <> "=ROUND(" Then
                            Target.Formula = "=ROUND(" & Mid(cell_formu_now, 2) & ", " & d & ")"
                        Else
                            Target.Formula = Left(cell_formu_now, InStrRev(cell_formu_now, ",") - 1) & ", " & d & ")"
                        End If
                    End If
                Else
                    If Left(cell_formu_now, 1) = "=" And Left(cell_formu_now, 7) = "=ROUND(" Then
                        Dim temp = Mid(cell_formu_now, 8)
                        Target.Formula = "=" & Left(temp, InStrRev(temp, ",") - 1)
                    End If
                End If

                Dim ar = Target.Address(RowAbsolute:=False, ColumnAbsolute:=False)


                '显示公式
                Dim cell_left = Target.Offset(0, -1)
                If Globals.Ribbons.Ribbon1.Chk2.Checked And Globals.Ribbons.Ribbon1.Chk3.Checked Then
                    cell_left.Formula = "=smart_formula(" & ar & ",0)"
                    Target.EntireColumn.AutoFit()
                    cell_left.EntireColumn.AutoFit()
                ElseIf Globals.Ribbons.Ribbon1.Chk2.Checked And Not Globals.Ribbons.Ribbon1.Chk3.Checked Then
                    cell_left.Formula = "=smart_formula(" & ar & ",1)"
                    Target.EntireColumn.AutoFit()
                    cell_left.EntireColumn.AutoFit()
                ElseIf Not Globals.Ribbons.Ribbon1.Chk2.Checked And Globals.Ribbons.Ribbon1.Chk3.Checked Then
                    cell_left.Formula = "=smart_formula(" & ar & ",2)"
                    Target.EntireColumn.AutoFit()
                    cell_left.EntireColumn.AutoFit()
                Else
                End If

                '命名单元格
                Dim cell_left2 = Target.Offset(0, -2)
                Dim n = cell_left2.Value
                Dim a = Target.Address
                Dim re = CreateObject("VBScript.RegExp") : re.Pattern = "[a-zA-Z]+(?=[0-9])"
                Dim matches = re.Execute(n)
                If matches.Count <> 0 Then
                    cell_left2.Value = Left(n, matches(0).Length) & "_" & Mid(n, matches(0).Length + 1)
                    n = cell_left2.Value
                End If
                If Globals.Ribbons.Ribbon1.Chk1.Checked And n IsNot Nothing Then
                    Try
                        xls.ActiveWorkbook.Names.Add(n, "=" & a)
                    Catch e As Exception
                        MsgBox(n & "所在单元格命名失败！" & vbCr & e.Message, vbExclamation)
                    End Try
                End If
                '量纲计算
                If Globals.Ribbons.Ribbon1.Chk5.Checked Then
                    Target.Offset(0, 1).Formula = unit_cal(xls.Range(ar))
                    Dim n_unit = n & "_unit"
                    Dim a_unit = Target.Offset(0, 1).Address
                    xls.ActiveWorkbook.Names.Add(n_unit, "=" & a_unit)
                End If

                xls.ScreenUpdating = True
            End If
        End If

        xls.EnableEvents = True
    End Sub


    Function unit_cal(c As Range)
        Dim var_f = c.Formula
        If var_f = "" Or Left(var_f, 1) <> "=" Then
            unit_cal = "<空>"
            Exit Function
        End If
        If Left(var_f, 7) = "=ROUND(" Then
            var_f = Mid(var_f, 8)
            var_f = "=" & Left(var_f, InStrRev(var_f, ",") - 1)
        End If


        Dim re = CreateObject("vbscript.regexp")
        re.Global = True

        For Each n1 In xls.ActiveWorkbook.Names
            If Right(n1.Name, 5) = "_unit" Then
                Dim v = Left(n1.Name, Len(n1.Name) - 5)
                Dim u = "(" & xls.Range(xls.ActiveWorkbook.Names.Item(n1.Name).RefersTo).Value & ")"

                re.Pattern = "[+\-*/^(,=]" & v & "(?=[+\-*/^),])"
                Dim matches = re.Execute(var_f)
                For Each i In matches
                    var_f = Replace(var_f, i.Value, Left(i.Value, 1) & u)
                Next

                re.Pattern = "[+\-*/^(,=]" & v & "$"
                matches = re.Execute(var_f)
                For Each i In matches
                    var_f = Replace(var_f, i.Value, Left(i.Value, 1) & u)
                Next

            End If
        Next
        Dim supported_funcs = Split(xls.ActiveWorkbook.Sheets("【settings】").Cells(1, 1).Text, vbLf)

        For Each func_ In supported_funcs
            re.Pattern = "[+\-*/^(,=]" & func_ & "(?=[+\-*/^(,])"
            Dim matches = re.Execute(var_f)
            For Each i In matches
                var_f = Replace(var_f, i.Value, Left(i.Value, 1))
            Next
        Next

        var_f = Replace(var_f, "~", "1") '单位1
        var_f = Replace(var_f, "SQRT", "sp.sqrt")
        var_f = Replace(var_f, "PI()", "3.14")
        For i = 2 To 9
            var_f = Replace(var_f, superscripts(i - 2), "^" & CStr(i))
        Next

        var_f = Mid(var_f, 2)
        Dim ws_out = ws.Exec("""" & py_location & "\py_unit_cal.exe""" & " """ & var_f & """").StdOut
        unit_cal = ws_out.ReadLine

        For i = 2 To 9
            unit_cal = Replace(unit_cal, "^" & CStr(i), superscripts(i - 2))
        Next
        If unit_cal = "1" Then unit_cal = "~"
    End Function


    Private Sub Application_WorkbookActivate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim ct = xls.ActiveWorkbook.Worksheets.Count
        For Each sh In Wb.Worksheets
            If sh.Name = "【settings】" Then Exit Sub
        Next

        xls.EnableEvents = False
        xls.ScreenUpdating = False
        Dim u_sheet = xls.ActiveWorkbook.Sheets.Add(After:=xls.ActiveWorkbook.Worksheets(ct))
        u_sheet.Name = "【settings】"
        u_sheet.Visible = 0

        Wb.Sheets("【settings】").Cells(1, 1).Value = Replace("MAX,MIN,SIN,COS,TAN,TREND,ABS", ",", vbLf)

        For i = 0 To 7
            Wb.Sheets("【settings】").Cells(1, i + 2).Value = superscripts(i)
        Next

        xls.ScreenUpdating = True
        xls.EnableEvents = True
    End Sub

End Class
