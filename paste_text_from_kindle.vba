Public Sub paste_text_from_kindle()
'
' #### DESCRIPTION:
'
' This VBA code is for use in pasting text that was copied from a Kindle book.
'
' Kindle makes alterations to the copied text, on the clipboard.
' * These alterations impair the usefulness of the text on the clipboard.
'
' This VBA-code fixes two such alterations:
' * Kindle adds bibliographic info at the end of the copied text.
'   * Example pasted-text is below.  Kindle added the empty-line and bibliographic info:
'       An integrated development environment (IDE) has the ability to greatly help or hinder development.
'
'       Amos, Brian. Hands-On RTOS with Microcontrollers: Building real-time embedded systems using FreeRTOS, STM32 MCUs, and SEGGER debug tools (p. 103). Packt Publishing. Kindle Edition.
' * Kindle replaces paragraph separtors with either hex:C2A020, or hex:20.
'   * Neither of those values are paragraph separators.
'     * Hex:20 is an ASCII space.
'     * Hex:C2A020 is a UTF8 non-breaking space, followed by an ASCII space
'   * So, the pasted text does not have line-feeds between paragraphs.
'
' This VBA-code pastes the clipboard, and:
' * Removes the added empty-line and bibliographic info
' * Replaces the Kindle paragraph-separators hex:C2A020 with two new-lines.
'   (The pasted paragraphs will be separated by an empty line.)
'
' This VBA-code does not alter the clipboard contents.
'
' Kindle may make other alterations to the copied text, on the clipboard.
' * Those alterations are not fixed by the present code.
' * For example, in a bulleted-list, Kindle removes the paragraph separator between list-items,
'   and it replaces the bullet-symbols with a space.
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
'
'
' #### TESTING:
'
' * This program was written for a particular Kindle book, and the program was tested with it:
'   * "Hands-On RTOS with Microcontrollers"
'
' * This program might not work with other Kindle books, if they make different alterations to the clipboard,
'   for copied text.
'
'
' #### BACKGROUND INFO:
'
' * REs
'   * https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference
'   * https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions
'
' Copyright (c) 2021 by Jim Yuill, under the MIT License
'

Dim MyData As DataObject
Dim strClip, strOut1, strOut2 As String

' Put clipboard in string variable
Set MyData = New DataObject
MyData.GetFromClipboard
strClip = MyData.GetText

' Find empty-line and last paragraph, and replace with "", using RE
Dim regEx As New RegExp
Dim strPattern As String: strPattern = "\r\n\r\n[^\n]*$"
Dim strReplace As String: strReplace = ""

' Multiline mode:  ^ and $ match the beginning and end of each line (instead of the beginning and end of the input string)
With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = False
    .Pattern = strPattern
End With

If regEx.test(strClip) Then
    strOut1 = regEx.Replace(strClip, strReplace)
Else
    strOut1 = strClip
End If

' Find Kindle's paragraph separators of the form hex:C2A020, and replace them with two new-lines, using RE
' * hex:C2A020
'   * C2A0 is UTF8 for non-breaking space
'   * 20 is ASCII space
' * In VBA regex, \u is the prefix for UTF8
strPattern = "[\u00A0][\x20]"
' In Kindle, paragraphs are separated by a blank line, so replace with two new-lines.
strReplace = vbNewLine & vbNewLine

With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = False
    .Pattern = strPattern
End With

If regEx.test(strOut1) Then
    strOut2 = regEx.Replace(strOut1, strReplace)
Else
    strOut2 = strOut1
End If

' Write string to document
Selection.TypeText strOut2

End Sub ' END: paste_text_from_kindle()
