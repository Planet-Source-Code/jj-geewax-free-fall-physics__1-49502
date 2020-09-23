VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Free Fall Exact"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   4095
      Left            =   1680
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton Command6 
         Caption         =   "Other Abouts"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit Program"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Go Back to Graphical Version"
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Calculations"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save List"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Drop"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Height Values"
      Height          =   4095
      Left            =   0
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
      Begin MSComDlg.CommonDialog CMD1 
         Left            =   840
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   ".txt"
         Filter          =   "*.txt | (*.txt) Text Files"
      End
      Begin VB.ListBox Points 
         Height          =   3765
         ItemData        =   "frmExact.frx":0000
         Left            =   120
         List            =   "frmExact.frx":0007
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin RichTextLib.RichTextBox RTX 
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmExact.frx":0013
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2375
      Left            =   3120
      Picture         =   "frmExact.frx":008E
      ScaleHeight     =   2340
      ScaleWidth      =   1905
      TabIndex        =   15
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Configurations"
      Height          =   1695
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1080
         TabIndex        =   5
         Text            =   ".01"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1080
         TabIndex        =   4
         Text            =   "9.8"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Text            =   "200"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Init Velocity"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Initial Velocity"
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m/s"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Delta T"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Number of Iterations"
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Acceleration"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Acceleration Due to Gravity  (positive in this case!!!)"
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m/s²"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Init Height"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Height the object is dropped from."
         Top             =   1320
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About the ""Exact"" Version"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtAbout 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmExact.frx":E950
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmExact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim InitVel As Double
Dim DeltaT As Double
Dim Acc As Double
Dim InitHeight As Double
Dim T As Double
Dim R As Double
T = 0# 'set our t to zero
Points.Clear 'clear the points list
InitVel = Round(Val(Text1(4).Text), 3) 'set the init vel
DeltaT = Round(Val(Text1(3).Text), 3) 'set the dt
Acc = Round(Val(Text1(2).Text), 3) 'set the acc
InitHeight = Round(Val(Text1(0).Text), 3) 'set the init height
Do 'start do loop
T = T + DeltaT 'first t will be .01 or so, whichever
R = Round(InitHeight - ((InitVel * T) + (0.5 * Acc * T * T)), 3)
    'do the calculations
Points.AddItem (R) 'add the point to the list
Loop Until R <= 0 'go until the object drops.
    'this could be a very long list!!!
End Sub

Private Sub Command2_Click()
On Error GoTo Err:
RTX.Text = vbNullString
CMD1.CancelError = True
CMD1.ShowSave

Call SaveListBox(CMD1.FileName, Points)
Err:
End Sub

Private Sub Command3_Click()
frmCalculations.Show vbModal
frmCalculations.Text1(2).Text = Text1(4).Text
frmCalculations.Text1(1).Text = Text1(2).Text
frmCalculations.Text1(0).Text = Text1(0).Text
End Sub

Private Sub Command4_Click()
frmExact.Hide
frmMain.Show

End Sub

Private Sub Command5_Click()
If MsgBox("Are you sure you want to exit altogether?", vbYesNo) = vbYes Then
End
End If
End Sub

Private Sub Command6_Click()
MsgBox "Just want to thank anyone for using the program" & vbNewLine & _
"If you have any questions, fixes, bug reports, or comments in general, " & _
"please feel free to email me at: " & vbNewLine & _
"jgeewax@standrews-de.org"

End Sub

Private Sub Text1_LostFocus(Index As Integer)
If IsNumeric(Text1(Index).Text) = False Then
    Text1(Index).Text = "0" 'we need something here!
End If
End Sub
