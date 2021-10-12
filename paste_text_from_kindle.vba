Public Sub paste_text_from_kindle()
'
' #### DESCRIPTION:
'
' This VBA code is for use in pasting text that was copied from a Kindle book.
'
' * Problem addressed:
'   * When using copy-and-paste from a Kindle book, Kindle adds bibliographic info at the end of the copied text.
'   * In the example pasted-text, below, Kindle added the blank-line and bibliographic info:
'       An integrated development environment (IDE) has the ability to greatly help or hinder development.
'
'       Amos, Brian. Hands-On RTOS with Microcontrollers: Building real-time embedded systems using FreeRTOS, STM32 MCUs, and SEGGER debug tools (p. 103). Packt Publishing. Kindle Edition.
'
' * What this VBA-code does:
'   * Pastes the clip-board, and removes the added blank-line and bibliographic info.
'
'
' #### INSTALLATION:
'
' * Word's VBA-system has to be configured so it can use the DataObject type:
'   * "Microsoft Forms... Object Library" needs to be added.
'   * The instructions for that are at this page, in the post by "Uwe G. G." on "June 13, 2018"
'     * https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_win10-mso_2016/mystery-compile-error-user-defined-type-not/b0c07a65-9f0c-43f1-a181-12c95db0ac8d
'
' * Word's VBA-system has to be configured so it can use regular-expressions:
'   * "Microsoft VBScript Regular Expressions" needs to be added.
'   * The instructions for that are at this page, in the post by "Automate This" on "Mar 20 '14", under "Step 1"
'     * https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'
' * Install this VBA code in your Word Normal.dotm file, and assign a hotkey to it.
'   * Instructions can be found on the Internet, e.g.,
'     * https://wordmvp.com/FAQs/MacrosVBA/CreateAMacro.htm
''
' #### BACKGROUND INFO:
'
' * REs
'   * https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference
'   * https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions
'
' Copyright (c) 2021 by Jim Yuill, under the MIT License
'

Dim MyData As DataObject
Dim strClip, strOut As String

' Put clipboard in string variable
Set MyData = New DataObject
MyData.GetFromClipboard
strClip = MyData.GetText

' Find last paragraph and replace with "", using RE
Dim regEx As New RegExp
Dim strPattern As String: strPattern = "\r\n[^\n]*$"
Dim strReplace As String: strReplace = ""

With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = False
    .Pattern = strPattern
End With

If regEx.test(strClip) Then
    strOut = regEx.Replace(strClip, strReplace)
Else
    strOut = strClip
End If

' Write string to document
Selection.TypeText strOut

End Sub ' END: paste_text_from_kindle()
