Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic

Friend Class frmDeleteTmpFiles
	Inherits System.Windows.Forms.Form
	
	
	Private Sub frmDeleteTmpFiles_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '		Dim strCommandLine As String
        '        Dim deleteFileSpec As String = ""
        '		Dim lPos As Integer
        '		Dim lPos2 As Integer
        '		Dim fsObject As Scripting.FileSystemObject
        '		Dim agedMinutes As Integer
        '		Dim configDir As String
        '		Dim fileNum As Short
        '		Dim textLine As String
        '		Dim dfsArr() As String
        '		Dim dfsExt() As String
        '		Dim agedMinArr() As Integer
        '		Dim dfsCount As Integer
        '		Dim lindex As Integer
        '		Dim oFiles As Scripting.Files
        '		Dim oFile As Scripting.File
        '		Dim oFolder As Scripting.Folder
        '        Dim fileExt As String = ""
        '		Dim timeNow As Date
        '		Dim timeCreated As Date
        '		Dim minutesDiff As Integer
        '		Dim foundDeleteDir As Boolean
        '        Dim timeDateStamp As Date

        '		On Error Resume Next

        '        ' write time date stamp
        '        timeDateStamp = DateAndTime.DateValue(

        '		' create object
        '		fsObject = New Scripting.FileSystemObject

        '        ReDim dfsArr(0 To 0)
        '        ReDim dfsExt(0 To 0)
        '        ReDim agedMinArr(0 To 0)

        '		' get command line
        '		strCommandLine = Trim(VB.Command())
        '		'strCommandLine = "C:\Inetpub\wwwroot\publish\PDF\*.pdf;30"

        '		' check command line
        '		If Len(strCommandLine) = 0 Then
        '			' read from cofig files for multiple delete paths
        '			configDir = getConfigDir

        '			Call writeToDebug("ConfigDir: " & configDir)

        '			' check for file, exit if doesn't exist
        '			If fsObject.FileExists(configDir) = False Then
        '				Me.Close()
        '				Exit Sub
        '			End If

        '			' open file and read in things to delete
        '			fileNum = FreeFile

        '			' dfscount to zero
        '			dfsCount = 0

        '			' open
        '			FileOpen(fileNum, configDir, OpenMode.Input)
        '			' Loop until end of file.
        '			Do While Not EOF(fileNum)
        '				' set default false
        '				foundDeleteDir = False
        '				' Read line into variable.
        '				textLine = LineInput(fileNum)

        '				' redim arrays
        '                ReDim Preserve dfsArr(dfsCount)
        '                ReDim Preserve agedMinArr(dfsCount)
        '                ReDim Preserve dfsExt(dfsCount)

        '                ' increase dfscount
        '                dfsCount += 1

        '				Call writeToDebug("Delete Line: " & textLine)

        '				' separate dfs and agedMinutes
        '				' get file spec to delete
        '				lPos = InStr(1, textLine, ";")

        '				' check lpos
        '				If lPos > 0 Then
        '					' get file spec to delete
        '					deleteFileSpec = Trim(VB.Left(textLine, lPos - 1))

        '					If InStr(1, deleteFileSpec, "%appPath%", CompareMethod.Text) > 0 Then
        '						' replace
        '						deleteFileSpec = Replace(deleteFileSpec, "%appPath%", appPath,  ,  , CompareMethod.Text)
        '					End If

        '					' check for 'falconPdfPublisher' in path before adding to array
        '					If InStr(1, deleteFileSpec, "falconPdfPublisher", CompareMethod.Text) = 0 Then foundDeleteDir = True
        '					If InStr(1, deleteFileSpec, "falconPlot", CompareMethod.Text) = 0 Then foundDeleteDir = True
        '					If InStr(1, deleteFileSpec, "falconWeb", CompareMethod.Text) = 0 Then foundDeleteDir = True

        '					If foundDeleteDir = False Then GoTo nextPath

        '					' get aged Minutes
        '					agedMinutes = CInt(VB.Right(textLine, Len(textLine) - lPos))

        '				End If

        '				' separate extensions from delete path
        '				lPos2 = InStrRev(deleteFileSpec, "\")

        '				' check
        '				If lPos2 > 0 Then
        '					' insert into array
        '					dfsArr(dfsCount) = VB.Left(deleteFileSpec, lPos2)
        '					dfsExt(dfsCount) = VB.Right(deleteFileSpec, Len(deleteFileSpec) - (lPos2 + 2))
        '				End If

        '				' insert into array
        '				agedMinArr(dfsCount) = agedMinutes

        '				' next path
        'nextPath: 

        '			Loop 

        '			FileClose(fileNum)

        '		Else
        '			' get file spec to delete
        '			lPos = InStr(1, strCommandLine, ";")

        '			' set to zero
        '			dfsCount = 0

        '			' check lpos
        '			If lPos > 0 Then
        '				' get file spec to delete
        '				deleteFileSpec = Trim(VB.Left(strCommandLine, lPos - 1))

        '				If InStr(1, deleteFileSpec, "%appPath%", CompareMethod.Text) > 0 Then
        '					' replace
        '					deleteFileSpec = Replace(deleteFileSpec, "%appPath%", appPath,  ,  , CompareMethod.Text)
        '				End If

        '				' set default false
        '				foundDeleteDir = False

        '				' check for 'falconPdfPublisher' in path before adding to array
        '				'            If InStr(1, deleteFileSpec, "falconPdfPublisher", vbTextCompare) = 0 Then GoTo exitOut

        '				' check for 'falconPdfPublisher' in path before adding to array
        '				If InStr(1, deleteFileSpec, "falconPdfPublisher", CompareMethod.Text) = 0 Then foundDeleteDir = True
        '				If InStr(1, deleteFileSpec, "falconPlot", CompareMethod.Text) = 0 Then foundDeleteDir = True
        '				If InStr(1, deleteFileSpec, "falconWeb", CompareMethod.Text) = 0 Then foundDeleteDir = True

        '				If foundDeleteDir = False Then GoTo exitOut


        '				' redim arrays
        '                ReDim Preserve dfsArr(dfsCount)
        '                ReDim Preserve agedMinArr(dfsCount)
        '                ReDim Preserve dfsExt(dfsCount)

        '                ' increase dfscount
        '                dfsCount += 1

        '				' get aged Minutes
        '				agedMinArr(dfsCount) = CInt(VB.Right(strCommandLine, Len(strCommandLine) - lPos))

        '				' separate extensions from delete path
        '				lPos2 = InStrRev(deleteFileSpec, "\")

        '				' check
        '				If lPos2 > 0 Then
        '					' insert into array
        '                    dfsArr(dfsCount) = deleteFileSpec.Substring(0, lPos2)

        '					dfsExt(dfsCount) = VB.Right(deleteFileSpec, Len(deleteFileSpec) - (lPos2 + 2))
        '				End If

        '			End If

        '		End If

        '		' loop through and delete files
        '        For lindex = 0 To dfsCount - 1
        '            ' check for folder
        '            Call writeToDebug("Checking for folder <" & dfsArr(lindex) & ">")
        '            ' check path before starting
        '            If fsObject.FolderExists(dfsArr(lindex)) = True Then
        '                ' set ofolder to Path folder
        '                oFolder = fsObject.GetFolder(dfsArr(lindex))

        '                ' check dirPath for files
        '                If oFolder.Files.Count > 0 Then
        '                    ' set to files in dirPath
        '                    oFiles = oFolder.Files

        '                    ' check for and add files
        '                    For Each oFile In oFiles
        '                        lPos = InStrRev(oFile.Name, ".")

        '                        ' check
        '                        If lPos > 0 Then
        '                            fileExt = VB.Right(oFile.Name, Len(oFile.Name) - lPos)
        '                        End If

        '                        Call writeToDebug("Filename: " & oFile.Name)

        '                        ' compare file extensions
        '                        If StrComp(dfsExt(lindex), fileExt, CompareMethod.Text) = 0 Then
        '                            ' found match for file extensions, check time
        '                            timeNow = Now
        '                            timeCreated = oFile.DateLastModified

        '                            ' check minutes different
        '                            'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        '                            minutesDiff = DateDiff(DateInterval.Minute, timeCreated, timeNow)

        '                            Call writeToDebug("Time difference: " & minutesDiff)
        '                            Call writeToDebug("Time to delete: " & agedMinArr(dfsCount))

        '                            ' check minutes different verse agedMinutes for delete or not
        '                            If minutesDiff >= agedMinArr(lindex) Then
        '                                Call writeToDebug("Deleting file <" & oFile.Path & ">")
        '                                ' delete file
        '                                Call fsObject.DeleteFile(oFile.Path, True)
        '                            End If

        '                        End If
        '                    Next oFile
        '                End If
        '            End If
        '        Next lindex

        'exitOut: 

        '		Me.Close()
		
	End Sub



End Class