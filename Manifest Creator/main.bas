Attribute VB_Name = "Module1"
'For hiding files
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Const FILE_ATTRIBUTE_HIDDEN = &H2

'For XP Controls
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Function StripPath(FilePath As String) As String
'*************************************************************
'Strips the filename out of a full path FilePath.
'Returns the stripped filename.
'*************************************************************

Dim X As Integer
Dim ct As Integer
    StripPath = FilePath
    X = InStr(FilePath, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, FilePath, "\")
    Loop
    If ct > 0 Then StripPath = Mid(FilePath, ct + 1)
    
End Function
Public Sub CreateManifest()
On Error GoTo errhandler

Dim CompanyName As String
Dim ProjectName As String
Dim SetHidden As Long

'Check to see if required fields are filled
If Form1.Text1.Text = "" Then MsgBox "Must enter an EXE name.", vbCritical: Exit Sub
If Form1.Text2.Text = "" Then MsgBox "Must enter a destination path.", vbCritical: Exit Sub

'Make sure the user included the extension in Text1, otherwise add one
lookingfor = ".exe"
lookingin = Form1.Text1.Text
mypos = InStr(lookingin, lookingfor)
If mypos Then 'If the extension was found, do nothing
Else 'Otherwise
Form1.Text1.Text = Form1.Text1.Text & ".exe" 'Add the extension to the text box
End If


'Fill in optional fields if they are blank
If Form1.Text3.Text = "" Then
CompanyName = "Company"
Else
CompanyName = Form1.Text3.Text
End If

If Form1.Text4.Text = "" Then
ProjectName = "Project"
Else
ProjectName = Form1.Text4.Text
End If

'Write the manifest file with the values from the textboxes
Open Form1.Text2.Text & Form1.Text1.Text & ".manifest" For Output As #1
Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
Print #1, "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">"
Print #1, "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""" & CompanyName & "." & ProjectName & "." & Form1.Text1.Text & """ type=""win32"" />"
Print #1, "<description>WindowsExecutable</description>"
Print #1, "<dependency>"
Print #1, "<dependentAssembly>"
Print #1, "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*"" />"
Print #1, "</dependentAssembly>"
Print #1, "</dependency>"
Print #1, "</assembly>"
Close #1

If Form1.Check1.Value = 1 Then 'If user wants to Create as Hidden
DestFile = Form1.Text2.Text & Form1.Text1.Text & ".manifest"
SetHidden = SetFileAttributes(DestFile, FILE_ATTRIBUTE_HIDDEN) 'Set the file property to hidden
End If


MsgBox "Manifest file created.", vbInformation

errhandler:
If Err.Number <> 0 Then 'If error occurs
MsgBox "Could not write to file.  It may already exist as a hidden file.", vbCritical
End If
End Sub
