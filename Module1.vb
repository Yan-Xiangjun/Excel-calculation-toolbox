﻿'Copyright 2023 阎相君
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
Module Module1
    Public xls As Excel.Application = Globals.ThisAddIn.Application
    Public superscripts()
    Public ws
    Public wdCharacter = 1
    Public wdParagraph = 4
    Public wdExtend = 1
    Public wdOMathInline = 1
    Public py_location
    Public my_function_xlam
End Module
