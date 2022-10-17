VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Star Signs"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   3690
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "starsigns.frx":0000
      Left            =   1920
      List            =   "starsigns.frx":0028
      TabIndex        =   25
      Text            =   "Combo2"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "starsigns.frx":0091
      Left            =   1920
      List            =   "starsigns.frx":00B9
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Star Sign Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1800
      TabIndex        =   18
      Top             =   720
      Width           =   1815
      Begin VB.CommandButton Command3 
         Caption         =   "Show Match"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Second Person`s Sign"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "First Person`s Sign"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Star Sign Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "&Aries"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show Profile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "&Pisces"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4200
         Width           =   975
      End
      Begin VB.OptionButton Option11 
         Caption         =   "A&Quarius"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3840
         Width           =   975
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Cap&ricorn"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Sa&gittarius"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "&Scorpio"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "L&ibra"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "&Virgo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "&Leo"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&Cancer"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Gemini"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Taurus"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   1800
      TabIndex        =   21
      Top             =   4440
      Width           =   1815
      Begin VB.CommandButton Command4 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1680
      Picture         =   "starsigns.frx":0122
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Star Signs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Deon van Zyl
Private Sub Command1_Click()

    End

End Sub

Private Sub Command2_Click()

    Form2.Show
    Form2.mnuView.Enabled = True
    Form2.Caption = Label4.Caption & " : " & Label5.Caption
    Form2.Image1.Picture = LoadPicture("starsigns" & "\" & Label4.Caption & "\" & Label4.Caption & ".gif")
    Form2.RichTextBox1.LoadFile ("starsigns" & "\" & Label4.Caption & ".txt")
    Form1.Visible = False
    

End Sub

Private Sub Command3_Click()

    Form2.Visible = True
    Form2.mnuView.Enabled = False
    On Error GoTo reverse
    Form2.RichTextBox1.LoadFile ("starsigns" & "\" & "match" & "\" & Combo1.Text & " with " & _
     Combo2.Text & ".txt")
     
    Form1.Visible = False
    Form2.Caption = "Match between " & Combo1.Text & " and " & Combo2.Text
     Exit Sub
     
reverse:

    Form2.RichTextBox1.LoadFile ("starsigns" & "\" & "match" & "\" & Combo2.Text & " with " & _
    Combo1.Text & ".txt")
     
    Form1.Visible = False
    Form2.Caption = "Match between " & Combo1.Text & " and " & Combo2.Text
    
    Exit Sub
    
End Sub

Private Sub Command4_Click()

    Option8.Value = True
    Label5.Caption = "Made by Deon on 2K/04/25"

End Sub




Private Sub Label4_Change()

    If Label4.Caption = "Made by Deon on 2K/04/25" Then
    Image1.Picture = LoadPicture("starsigns" & "\" & "Scorpio" & ".gif")
    Else
    Image1.Picture = LoadPicture("starsigns" & "\" & Label4.Caption & ".gif")
    End If

End Sub



Private Sub Option1_Click()

    Label4.Caption = "Aries"
    Label5.Caption = "21 March - 19 April"

End Sub

Private Sub Option10_Click()

    Label4.Caption = "Capricorn"
    Label5.Caption = "22 December - 19 January"

End Sub

Private Sub Option11_Click()

    Label4.Caption = "Aquarius"
    Label5.Caption = "20 January - 18 February"

End Sub

Private Sub Option12_Click()

    Label4.Caption = "Pisces"
    Label5.Caption = "19 February - 20 March"

End Sub

Private Sub Option2_Click()

    Label4.Caption = "Taurus"
    Label5.Caption = "20 April - 20 May"

End Sub

Private Sub Option3_Click()

    Label4.Caption = "Gemini"
    Label5.Caption = "21 May - 20 June"

End Sub

Private Sub Option4_Click(Index As Integer)

    Label4.Caption = "Cancer"
    Label5.Caption = "21 June - 22 July"

End Sub

Private Sub Option5_Click()

    Label4.Caption = "Leo"
    Label5.Caption = "23 July - 22 August"

End Sub

Private Sub Option6_Click()

    Label4.Caption = "Virgo"
    Label5.Caption = "23 August - 22 September"

End Sub

Private Sub Option7_Click()

    Label4.Caption = "Libra"
    Label5.Caption = "23 September - 22 October"

End Sub

Private Sub Option8_Click()

    Label4.Caption = "Scorpio"
    Label5.Caption = "23 October - 21 November"

End Sub

Private Sub Option9_Click()

    Label4.Caption = "Sagittarius"
    Label5.Caption = "22 November - 21 December"

End Sub
