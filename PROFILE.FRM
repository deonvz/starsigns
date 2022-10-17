VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "Starsign Profile"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"profile.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   120
      Picture         =   "profile.frx":00EC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2850
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuPicture 
         Caption         =   "Picture"
      End
      Begin VB.Menu mnuplanet 
         Caption         =   "Planet"
      End
      Begin VB.Menu mnustone 
         Caption         =   "Stone"
      End
      Begin VB.Menu mnusymbol 
         Caption         =   "Symbol Of"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()

     If Val(Form2.ScaleWidth) < 3900 Or Val(Form2.ScaleHeight) < 1900 Then
    Form2.Width = 3900
    Form2.Height = 5865
    Form2.RichTextBox1.Width = Form2.ScaleWidth - 200
    Form2.RichTextBox1.Height = Form2.ScaleHeight - 1200
    Form2.Image1.Left = Val(Form2.Width) / 2 - 1400
    Else
    Form2.RichTextBox1.Width = Form2.ScaleWidth - 200
    Form2.RichTextBox1.Height = Form2.ScaleHeight - 1200
    Form2.Image1.Left = Val(Form2.Width) / 2 - 1400
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Form1.Visible = True
    Unload Me

End Sub

Private Sub mnuexit_Click()

    Form1.Visible = True
    Unload Me

End Sub

Private Sub mnuPicture_Click()

    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & Form1.Label4.Caption & ".gif")

End Sub

Private Sub mnuplanet_Click()

    On Error GoTo DriveError
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "planet" & ".gif")
    Exit Sub

DriveError:
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "planet" & ".jpeg")
     Exit Sub


End Sub

Private Sub mnuprint_Click()

    On Error GoTo printerror
    Form2.PrintForm
    Exit Sub
    
printerror:
    
    MsgBox "No Printer found Please check if printer is plugged in and switched on", vbCritical
    Exit Sub

End Sub

Private Sub mnustone_Click()

    On Error GoTo DriveError
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "stone" & ".gif")
    Exit Sub
    
DriveError:
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "stone" & ".jpeg")
     Exit Sub

End Sub

Private Sub mnusymbol_Click()

    On Error GoTo DriveError
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "symbol" & ".gif")
    Exit Sub
    
DriveError:
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Form1.Label4.Caption & "\" & "symbol" & ".jpeg")
     Exit Sub
End Sub
