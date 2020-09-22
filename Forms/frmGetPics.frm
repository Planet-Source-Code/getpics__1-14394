VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGetPics 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox txtURL 
      Height          =   315
      ItemData        =   "frmGetPics.frx":0000
      Left            =   1080
      List            =   "frmGetPics.frx":0010
      TabIndex        =   3
      Top             =   60
      Width           =   4875
   End
   Begin InetCtlsObjects.Inet inetGetPics 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton btnGetPics 
      Caption         =   "&GetPics"
      Height          =   375
      Left            =   5940
      TabIndex        =   2
      Top             =   60
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwOnlinePics 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5636
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Rechts
      Caption         =   "URL:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "frmGetPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGetPics_Click()

    'http://www.jensschlegel.de/Site/AboutMe/Bilder/Bilder.htm
    
    Dim strResult As String
    Dim intCounter As Integer
    Dim itmPic As ListItem
    
    lvwOnlinePics.ListItems.Clear
    If Left(txtURL.Text, 7) <> "http://" Then
        txtURL.Text = "http://" & txtURL.Text
    End If
    If txtURL.Text <> "" Then
        strResult = inetGetPics.OpenURL(txtURL.Text, icString)
        If strResult <> "" Then
            fldPics = ParsePageForPics(strResult)
            If UBound(fldPics) > 0 Then
                For intCounter = 1 To UBound(fldPics)
                    If fldPics(intCounter) <> "" Then
                        Set itmPic = lvwOnlinePics.ListItems.Add(, , Str(intCounter))
                        itmPic.SubItems(1) = fldPics(intCounter)
                    End If
                Next
            End If
        End If
    End If

End Sub

Private Sub Form_Load()

    lvwOnlinePics.ColumnHeaders(1).Width = (lvwOnlinePics.Width / 100) * 15
    lvwOnlinePics.ColumnHeaders(2).Width = (lvwOnlinePics.Width / 100) * 84

End Sub


Private Sub lvwOnlinePics_DblClick()
    
    Dim fldResult() As Byte
    
    strImageName = GetImageName(lvwOnlinePics.SelectedItem.SubItems(1))
    If strImageName <> "" Then
        fldResult = inetGetPics.OpenURL(lvwOnlinePics.SelectedItem.SubItems(1), icByteArray)
        If UBound(fldResult) > 0 Then
            Open App.Path & "\" & strImageName For Binary Access Write As #1
            Put #1, , fldResult()
            Close #1
            Load frmDisplay
            frmDisplay.Show
        End If
    End If
   
End Sub
