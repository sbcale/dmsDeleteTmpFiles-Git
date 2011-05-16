Module testAddMod
    Public Function dmsV7Util_checkForInvalidCharsForFilenames(ByVal sNameToCheck As String) As Boolean
        Dim iPos As Integer

        dmsV7Util_checkForInvalidCharsForFilenames = False

        iPos = InStr(sNameToCheck, ":", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "\", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "/", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "*", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "?", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, Chr(34), CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "<", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, ">", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        iPos = InStr(sNameToCheck, "|", CompareMethod.Text)
        If iPos > 0 Then Exit Function

        dmsV7Util_checkForInvalidCharsForFilenames = True
    End Function

    Function RmNull(ByVal CharStr As String) As String

        ' This function will truncate 'CharStr' at the first NULL character found.

        ' Local variable declarations.

        Dim NullPos As Integer
        Dim tmpCharStr As String

        ' Find the first NULL character.

        tmpCharStr = CharStr

        NullPos = InStr(1, tmpCharStr, vbNullChar)

        ' Check 'NullPosition'.

        If NullPos < 1 Then tmpCharStr = Trim$(tmpCharStr)
        If NullPos = 1 Then tmpCharStr = ""
        If NullPos > 1 Then tmpCharStr = Trim$(Left$(tmpCharStr, NullPos - 1))

        ' The function value is never checked by the calling routine.  This should be a subroutine.

        RmNull = tmpCharStr

    End Function

End Module
