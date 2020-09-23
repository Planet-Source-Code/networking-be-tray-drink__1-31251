VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Not to be visible at any time..."
   ClientHeight    =   540
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrIconUpdate 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuDrink 
      Caption         =   "Drink"
      Begin VB.Menu mnuFill 
         Caption         =   "Fill her up!"
      End
      Begin VB.Menu mnuDegreeOfThirst 
         Caption         =   "Degree of thirst"
         Begin VB.Menu mnuDegree 
            Caption         =   "Very thirsty"
            Index           =   1
         End
         Begin VB.Menu mnuDegree 
            Caption         =   "Thirsty"
            Index           =   2
         End
         Begin VB.Menu mnuDegree 
            Caption         =   "Why not"
            Index           =   3
         End
         Begin VB.Menu mnuDegree 
            Caption         =   "Not thirsty"
            Index           =   4
         End
         Begin VB.Menu mnuDegree 
            Caption         =   "I'm full"
            Index           =   5
         End
      End
      Begin VB.Menu mnuChoose 
         Caption         =   "Choose drink"
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Traydrink - by cakkie
' just some vars to keep track of stuff
Dim DrinkSpeed As Long
Dim iCount As Long
Dim CurrentIcon As Integer
Dim CurrentDrink As String

Public Sub StartDrink(DrinkName As String)

    ' called by tray menu, load a new drink
    iCount = 0
    CurrentIcon = 100
    CurrentDrink = DrinkName
    
    AddToTray Me, "Drinking " & CurrentDrink, LoadPicture(App.Path & "\icons\" & CurrentDrink & "\" & CurrentIcon & ".ico"), True

End Sub

Private Sub Form_Load()

    ' get initial settings from registry (if exist), and load the drink
    iCount = 0
    CurrentIcon = 100
    
    CurrentDrink = GetSetting("Tray bar", "config", "drink", "beer")
    DrinkSpeed = GetSetting("Tray bar", "config", "speed", "2")
    
    AddToTray Me, "Drinking " & CurrentDrink, LoadPicture(App.Path & "\icons\" & CurrentDrink & "\" & CurrentIcon & ".ico")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    ' when button is pressed, this event will be raised via callback
    Dim Message As Long
    
    On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    
    Select Case Message
    Case WM_RBUTTONUP
            PopupMenu mnuDrink
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' cleanup and save
    RemoveFromTray
    SaveSetting "Tray bar", "config", "drink", CurrentDrink
    SaveSetting "Tray bar", "config", "speed", DrinkSpeed
    
End Sub

Private Sub mnuAbout_Click()

    ' need to say more?
    MsgBox "TrayBar - By Cakkie"

End Sub

Private Sub mnuChoose_Click()

    ' show form for selecting drink
    frmSelect.Show

End Sub

Private Sub mnuDegree_Click(Index As Integer)

    ' this will set how fast the drink gets empty
    DrinkSpeed = (Index ^ 2) * 2

End Sub

Private Sub mnuExit_Click()

    ' quit
    If vbYes = MsgBox("Are you sure you don't want another drink?", vbYesNo, "Quit") Then
        Unload Me
    End If

End Sub

Private Sub mnuFill_Click()

    ' refill drink
    CurrentIcon = 100
    iCount = 0
    AddToTray Me, "Drinking " & CurrentDrink, LoadPicture(App.Path & "\icons\" & CurrentDrink & "\0.ico"), True

End Sub

Private Sub tmrIconUpdate_Timer()
    
    iCount = iCount + 1
    If iCount = DrinkSpeed Then
        ' advance to next stage
        CurrentIcon = CurrentIcon - 25
        iCount = 0
    End If

    If CurrentIcon < 0 Then
        ' flash warning that drink is empty
        If iCount Mod 2 = 0 Then
            AddToTray Me, "Warning, drink empty", LoadPicture(App.Path & "\icons\" & CurrentDrink & "\0.ico"), True
        Else
            AddToTray Me, "Warning, drink empty", LoadPicture(App.Path & "\icons\" & CurrentDrink & "\!.ico"), True
        End If
    Else
        ' show drink
        AddToTray Me, "Drinking " & CurrentDrink, LoadPicture(App.Path & "\icons\" & CurrentDrink & "\" & CurrentIcon & ".ico"), True
    End If

End Sub
