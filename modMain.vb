Option Explicit On

Imports System.IO

Module modMain

    Friend DebugFile As String
    Friend binPath As String
    Friend appPath As String
    Friend configDir As String

    Friend Sub Main()
        ' start process
        Call processTmpDelete()

        Call writeToDebug("*******************************************************************")
        Call writeToDebug("* Finished running dmsDeleteTmpFiles: " & Now)
        Call writeToDebug("*******************************************************************")

    End Sub

    Friend Sub processTmpDelete()
        Dim strCommandLine As String
        Dim lPos As Integer
        Dim lPos2 As Integer
        Dim configDir As String
        Dim fileNum As Short
        Dim textLine As String
        Dim dfsArr() As String
        Dim dfsExt() As String
        Dim agedMinArr() As Integer
        Dim dfsCount As Integer
        Dim lindex As Integer
        Dim oFolder As Directory
        Dim timeNow As Date
        Dim timeCreated As Date
        Dim minutesDiff As Integer
        Dim iHours As Integer
        Dim iMinutes As Integer
        Dim fileExt As String = ""
        Dim deleteFileSpec As String = ""
        Dim foundDeleteDir As Boolean = False

        On Error Resume Next

        ' init arrays
        ReDim dfsArr(0 To 0)
        ReDim dfsExt(0 To 0)
        ReDim agedMinArr(0 To 0)

        ' get command line
        strCommandLine = My.Application.CommandLineArgs(0).ToString.Trim

        ' hardcoded temp SBC Merge test take 2
        'strCommandLine = "C:\Inetpub\wwwroot\publish\PDF\*.pdf;30"

        ' check command line
        If strCommandLine.Length = 0 Then
            ' read from cofig files for multiple delete paths
            configDir = getConfigDir()

            Call writeToDebug("ConfigDir: " & configDir)

            ' check for file, exit if doesn't exist
            If File.Exists(configDir) = False Then
                Call writeToDebug("** ConfigDir: " & configDir & " does not exist")
                ' exit
                Application.Exit()
                Exit Sub
            End If

            ' open file and read in things to delete
            fileNum = FreeFile()

            ' dfscount to zero
            dfsCount = 0

            'Dim objReader As New StreamReader(configDir)

            Using objReader As New StreamReader(configDir, True)
                ' loop to end of file
                While Not objReader.EndOfStream
                    ' set default false
                    foundDeleteDir = False
                    ' read current line
                    textLine = objReader.ReadLine.Trim

                    ' redim arrays
                    ReDim Preserve dfsArr(dfsCount)
                    ReDim Preserve agedMinArr(dfsCount)
                    ReDim Preserve dfsExt(dfsCount)

                    Call writeToDebug("Delete Line: " & textLine)

                    ' separate dfs and agedMinutes
                    ' get file spec to delete
                    lPos = InStr(textLine, ";")

                    ' check lpos
                    If lPos > 0 Then
                        ' get file spec to delete
                        deleteFileSpec = Trim(Left(textLine, lPos - 1))

                        If InStr(deleteFileSpec, "%appPath%", CompareMethod.Text) > 0 Then
                            ' replace
                            'deleteFileSpec = Replace(deleteFileSpec, "%appPath%", appPath, , , CompareMethod.Text)
                            deleteFileSpec = deleteFileSpec.Replace("%appPath%", appPath).Trim
                        End If

                        ' check for 'falconPdfPublisher' in path before adding to array
                        If InStr(deleteFileSpec, "falconPdfPublisher", CompareMethod.Text) > 0 Then foundDeleteDir = True
                        If InStr(deleteFileSpec, "falconPlot", CompareMethod.Text) > 0 Then foundDeleteDir = True
                        If InStr(deleteFileSpec, "falconWeb", CompareMethod.Text) > 0 Then foundDeleteDir = True

                    End If

                    ' check 
                    If foundDeleteDir = True Then
                        ' insert into array
                        agedMinArr(dfsCount) = CInt(Right(textLine, textLine.Length - lPos))

                        ' separate extensions from delete path
                        lPos2 = InStrRev(deleteFileSpec, "\")

                        ' check
                        If lPos2 > 0 Then
                            ' insert into array
                            dfsArr(dfsCount) = Left(deleteFileSpec, lPos2)
                            dfsExt(dfsCount) = Right(deleteFileSpec, deleteFileSpec.Length - (lPos2 + 2))
                        End If

                        ' increase dfscount
                        dfsCount += 1
                    End If

                End While

                ' close file
                objReader.Close()

            End Using

        Else
            ' get file spec to delete
            lPos = InStr(strCommandLine, ";")

            ' set to zero
            dfsCount = 0

            ' check lpos
            If lPos > 0 Then
                ' get file spec to delete
                deleteFileSpec = Trim(Left(strCommandLine, lPos - 1))

                If InStr(deleteFileSpec, "%appPath%", CompareMethod.Text) > 0 Then
                    ' replace
                    deleteFileSpec = deleteFileSpec.Replace("%appPath%", appPath).Trim
                End If

                ' set default false
                foundDeleteDir = False

                ' check for 'falconPdfPublisher' in path before adding to array
                If InStr(deleteFileSpec, "falconPdfPublisher", CompareMethod.Text) > 0 Then foundDeleteDir = True
                If InStr(deleteFileSpec, "falconPlot", CompareMethod.Text) > 0 Then foundDeleteDir = True
                If InStr(deleteFileSpec, "falconWeb", CompareMethod.Text) > 0 Then foundDeleteDir = True
                ' check for dir in path
                If foundDeleteDir = False Then
                    ' exit
                    Application.Exit()
                    ' exit
                    Exit Sub
                End If

                ' redim arrays
                ReDim Preserve dfsArr(dfsCount)
                ReDim Preserve agedMinArr(dfsCount)
                ReDim Preserve dfsExt(dfsCount)

                ' get aged Minutes
                agedMinArr(dfsCount) = CInt(Right(strCommandLine, strCommandLine.Length - lPos))

                ' separate extensions from delete path
                lPos2 = InStrRev(deleteFileSpec, "\")

                ' check
                If lPos2 > 0 Then
                    ' insert into array
                    dfsArr(dfsCount) = deleteFileSpec.Substring(0, lPos2)

                    dfsExt(dfsCount) = Right(deleteFileSpec, deleteFileSpec.Length - (lPos2 + 2))
                End If

            End If

            ' increase dfscount
            dfsCount += 1

        End If

        ' loop through and delete files
        For lindex = 0 To dfsExt.Length - 1
            ' check for folder
            Call writeToDebug("Checking for folder <" & dfsArr(lindex) & ">")
            ' check path before starting
            If Directory.Exists(dfsArr(lindex)) = True Then
               
                ' check dirPath for files
                'If Directory.GetDirectories(dfsArr(lindex)).Length > 0 Then
                If Directory.GetFiles(dfsArr(lindex)).Length > 0 Then

                    ' check for and add files
                    For Each filename As String In Directory.GetFiles(dfsArr(lindex))
                        ' extract file extension
                        fileExt = Path.GetExtension(filename.Trim)
                        ' remove .
                        fileExt = Right(fileExt, fileExt.Length - 1)

                        Call writeToDebug("Filename: " & filename.Trim)

                        ' compare file extensions
                        If StrComp(dfsExt(lindex), fileExt, CompareMethod.Text) = 0 Or StrComp(dfsExt(lindex), "*", CompareMethod.Binary) = 0 Then
                            ' found match for file extensions, check time
                            timeNow = Now
                            timeCreated = File.GetLastWriteTime(filename)


                            ' check minutes different
                            'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
                            minutesDiff = DateDiff(DateInterval.Minute, timeCreated, timeNow)

                            ' get minutes in hour/min
                            iHours = Int(minutesDiff / 60)
                            iMinutes = minutesDiff Mod 60

                            If iHours <> 0 Then
                                ' minutes and hours
                                Call writeToDebug("Time difference: " & iHours & " hours " & iMinutes & " minutes")
                            Else
                                ' only minutes
                                Call writeToDebug("Time difference: " & iMinutes & " minutes")
                            End If

                            Call writeToDebug("Time to delete: " & agedMinArr(lindex).ToString & " minutes")

                            ' check minutes different verse agedMinutes for delete or not
                            If minutesDiff >= agedMinArr(lindex) Then
                                Call writeToDebug("** Deleting file <" & filename.Trim & ">")

                                ' clear errors
                                Err.Clear()

                                ' delete file
                                Call File.Delete(filename.Trim)

                                ' check for error
                                If Err.Number <> 0 Then
                                    ' write to log
                                    Call writeToDebug("** Error deleting file: <" & filename.Trim & ">" & " Error: " & Err.Description)
                                    ' clear errors
                                    Err.Clear()
                                End If

                            Else ' minutesDiff
                                Call writeToDebug("* Not Deleting file <" & filename.Trim & ">. Not old enough.")
                            End If
                        Else ' extension match
                            Call writeToDebug("* Not Deleting file <" & filename.Trim & ">. Not a file extension match.")
                        End If
                    Next
                End If
            End If
        Next lindex

        ' exit
        Application.Exit()
    End Sub

    Private Function getConfigDir() As String
        Dim configDir As String

        ' default empty
        getConfigDir = ""

        ' get apppath
        binPath = My.Application.Info.DirectoryPath.Trim

        ' hardcode temp
        '    binPath = "C:\Web Applications\falconPdfPublisher\bin"

        ' check for trailing slash and add if not there
        If StrComp(binPath.Substring(binPath.Length), "\") <> 0 Then binPath = binPath & "\"

        ' compose paths
        configDir = binPath.Replace("\bin\", "\app_data\cfg\").Trim
        appPath = binPath.Replace("\bin\", "").Trim
        DebugFile = binPath.Replace("\bin\", "\app_data\log\").Trim & "dmsTempDelete.log"

        ' write time date stamp
        Call writeToDebug("")
        Call writeToDebug("************************************************************")
        Call writeToDebug("* Running dmsDeleteTmpFiles: " & Now)
        Call writeToDebug("************************************************************")

        ' return
        getConfigDir = configDir & "dmsDeleteTmpFiles.cfg"

    End Function

    Sub writeToDebug(ByRef tmp As String)
        ' delcares
        Dim objWriter As New StreamWriter(DebugFile, True)

        objWriter.WriteLine(tmp)
        objWriter.Flush()
        objWriter.Close()

    End Sub


    ' This function will truncate 'CharStr' at the first NULL character found.
    ' Local variable declarations.
    Sub RemoveNull(ByRef CharStr As String)
        Dim NullPos As Short
        ' Find the first NULL character.
        NullPos = InStr(CharStr, vbNullChar)
        ' Check 'NullPosition'.
        If NullPos < 1 Then CharStr = Trim(CharStr)
        If NullPos = 1 Then CharStr = ""
        If NullPos > 1 Then CharStr = Trim(Left(CharStr, NullPos - 1))
    End Sub
End Module