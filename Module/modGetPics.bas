Attribute VB_Name = "modGetPics"
Public fldPics() As String
Public strImageName As String

Sub Main()

    Load frmGetPics
    frmGetPics.Show

End Sub

Public Function ParsePageForPics(ByVal strPage As String) As Variant

    Dim intTagStartPos, intTagEndPos, intPicCounter, intLastRelativeSign As Integer
    Dim strURL, strAbsolutePicture As String
    
    strURL = GetURL
    intPicCounter = 1
    intTagStartPos = InStr(strPage, "<img src=" & Chr$(34))
    If intTagStartPos > 0 Then
        Do
            intTagEndPos = InStr(intTagStartPos + 11, strPage, Chr$(34))
            If intTagEndPos > 0 Then
                ReDim Preserve fldPics(intPicCounter)
                fldPics(intPicCounter) = Mid(strPage, intTagStartPos + 10, (intTagEndPos - 2) - (intTagStartPos + 8))
                intPicCounter = intPicCounter + 1
            End If
            intTagStartPos = InStr(intTagEndPos, strPage, "<img src=" & Chr$(34))
        Loop Until intTagStartPos = 0
    End If
    If intPicCounter <= 1 Then
        MsgBox "No Pics Available!"
        Exit Function
    End If
    For intCounter = 1 To UBound(fldPics)
        If InStr(fldPics(intCounter), "../") > 0 Then
            intLastRelativeSign = InStrRev(fldPics(intCounter), "../") + 3
            strAbsolutePicture = strURL & Mid(fldPics(intCounter), intLastRelativeSign, Len(fldPics(intCounter)) - intLastRelativeSign + 1)
            fldPics(intCounter) = strAbsolutePicture
        End If
        If Left(fldPics(intCounter), 7) <> "http://" Then
            fldPics(intCounter) = strURL & fldPics(intCounter)
        End If
    Next intCounter
    
    ParsePageForPics = fldPics

End Function

Public Function GetURL() As String

    Dim intLastSlashPos As Integer
    
    intLastSlashPos = InStr(8, frmGetPics.txtURL, "/")
    If intLastSlashPos <> 0 Then
        GetURL = Left(frmGetPics.txtURL, intLastSlashPos)
    Else
        If Right(frmGetPics.txtURL, 1) <> "/" Then
            GetURL = frmGetPics.txtURL & "/"
        Else
            GetURL = frmGetPics.txtURL
        End If
    End If
    
End Function

Public Function GetImageName(strPicture As String) As String

    Dim intLastSlashPos As Integer
    
    intLastSlashPos = InStrRev(strPicture, "/") + 1
    GetImageName = Mid(strPicture, intLastSlashPos, Len(strPicture) - intLastSlashPos + 1)

End Function
