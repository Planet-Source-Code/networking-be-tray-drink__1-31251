VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select drink"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order this drink"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox lstDrink 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOrder_Click()

    ' set new drink
    If lstDrink.ListIndex = -1 Then
        MsgBox "You must select a drink in order to get one, duh!"
    Else
        frmMain.StartDrink lstDrink.Text
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    ' list folders (drinks)
    Dim fso, subfolder, folder
    
    Set fso = CreateObject("Scripting.Filesystemobject")
    Set folder = fso.getfolder(App.Path & "\icons\")
    
    For Each subfolder In folder.subfolders
        lstDrink.AddItem subfolder.Name
    Next subfolder

End Sub

Private Sub lstDrink_DblClick()

    If lstDrink.ListIndex <> -1 Then cmdOrder_Click

End Sub
