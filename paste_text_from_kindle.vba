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
'
' * Kindle replaces paragraph separtors with either hex:C2A020, or hex:20.
'   * Neither of those values are paragraph separators.
'     * Hex:20 is an ASCII space.
'     * Hex:C2A020 is a UTF8 non-breaking space, followed by an ASCII space
'   * So, the pasted text does not have line-feeds between paragraphs.
'
' This VBA-code pastes the clipboard, and:
' * Removes the added empty-line and bibliographic info
' * Replaces the Kindle paragraph-separators hex:C2A020 with two new-lines.
'   * These pasted paragraphs will be separated by an empty line.
' * If Kindle used Hex:20 (ASCII space) as the paragraph separator, then the pasted text's
'   paragraphs will not start on a new line nor be separated by an empty line.
' * Sets the pasted text to italics as specified in the configuration variable SET_ITALICS_ON
'   * More info about SET_ITALICS_ON is below.
'
' This VBA-code does not alter the clipboard contents.
'
' Kindle may make other alterations to the copied text, on the clipboard.
' * Those alterations are not fixed by the present code.
' * For example, in a bulleted-list, Kindle removes the paragraph separator between list-items,
'   and it replaces the bullet-symbols with a space.
'
'
'
' #### INSTALLATION:
'
' * Install the present VBA code in a Word document or template
'   * For example, the VBA code can be installed in your Normal.dotm file,
'     and a hotkey can be assigned to the VBA code.
'     * This will make the macro available whenever you use Word on your computer.
'   * Installation instructions can be found on the Internet, e.g.,
'     * An overview of Normal.dotm, and where to find it:
'       * See the section, "Normal.dotm - the pan-global template - the granddaddy of all document templates"
'       * http://www.addbalance.com/usersguide/templates.htm
'     * What do I do with macros sent to me by other users to help me out?
'       * https://wordmvp.com/FAQs/MacrosVBA/CreateAMacro.htm
'     * How to assign a Word command or macro to a hot-key?
'       * https://wordmvp.com/FAQs/Customization/AsgnCmdOrMacroToHotkey.htm
'
' * Configure the Word document or tempalate in which the VBA code is installed
'   * Open the Word document or tempalate
'     * To open a template, right click on it, and select "Open".  Do not double-click on the template.
'   * Open the VBA editor, using Alt+F11
'   * Follow the instructions below to set-up access to the needed VBA functions
'     * Configure access to the DataObject type:
'       * "Microsoft Forms... Object Library" is needed.
'       * The instructions for that are at this page, in the post by "Uwe G. G." on "June 13, 2018"
'         * https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_win10-mso_2016/mystery-compile-error-user-defined-type-not/b0c07a65-9f0c-43f1-a181-12c95db0ac8d
'     * Configure access to regular-expressions:
'       * "Microsoft VBScript Regular Expressions" is needed.
'       * The instructions for that are at this page, in the post by "Automate This" on "Mar 20 '14", under "Step 1"
'         * https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'
'
' #### TESTING:
'
' * This program has been tested with two Kindle books.
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
' Copyright (c) 2021-2022 by Jim Yuill, under the MIT License
'

' Configuration variable:
' * To paste text as italics, SET_ITALICS_ON should be set equal to "True"
' * To not paste text as italics, SET_ITALICS_ON should be set equal to "False"
Const SET_ITALICS_ON As String = "False"


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

' Info on setting italics
' * http://vbcity.com/forums/t/131307.aspx
'   * See last post
' * https://stackoverflow.com/questions/31558157/vba-word-selection-typetext-changing-font
'   * See post by Paul Ogilvie
With Selection
    .InsertAfter (strOut2)
    If SET_ITALICS_ON = "True" Then
        .Font.Italic = True
    End If
End With


End Sub ' END: paste_text_from_kindle()
