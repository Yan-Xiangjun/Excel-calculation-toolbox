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
Imports System.Windows.Forms

Public Class excel_functions

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        xls.ActiveWorkbook.Sheets("【settings】").Cells(1, 1).Value = Replace(tb1.Text, vbCrLf, vbLf)
        Close()
    End Sub

    Private Sub excel_functions_Load(sender As Object, e As EventArgs) Handles Me.Load
        tb1.Text = Replace(xls.ActiveWorkbook.Sheets("【settings】").Cells(1, 1).Text, vbLf, vbCrLf)
    End Sub


End Class