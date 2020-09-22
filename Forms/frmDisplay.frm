VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Image imgPic 
      Height          =   4335
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "DoubleClick to close!!!"
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    imgPic.Picture = LoadPicture(App.Path & "\" & strImageName)
    imgPic.Refresh
    
End Sub


Private Sub imgPic_DblClick()

    Kill App.Path & "\" & strImageName
    Unload frmDisplay

End Sub
