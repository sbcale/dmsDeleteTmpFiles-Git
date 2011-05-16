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
End Module
