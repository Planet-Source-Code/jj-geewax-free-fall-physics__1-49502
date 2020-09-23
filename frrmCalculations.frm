VERSION 5.00
Begin VB.Form frmCalculations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculations"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Other"
      Height          =   1335
      Left            =   2640
      TabIndex        =   25
      Top             =   2880
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Set Precision"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Create Report"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close Window"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Precision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Initial Velocity"
      Height          =   1335
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   2175
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "H0 at:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Initial Height"
      Height          =   1215
      Left            =   2640
      TabIndex        =   17
      Top             =   1560
      Width           =   2175
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "H0 at:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instantaneous Velocity"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "At time:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "When will it hit the ground?"
      Height          =   1335
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate!"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Values"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
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
         Left            =   720
         TabIndex        =   3
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
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Text            =   "9.8"
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
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Text            =   "200"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Init Vel"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Initial Velocity"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m/s"
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Acc."
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Acceleration Due to Gravity  (positive in this case!!!)"
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m/s²"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Height the object is dropped from."
         Top             =   960
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Command1_Click()
    Label1.Caption = Round(FindTimeHitsGround(Text1(1).Text, Text1(2).Text, Text1(0).Text), Int(Precision.Caption)) & " seconds"
End Sub
Private Sub Command2_Click()
    frmCalculations.Hide
    frmMain.Show
End Sub
Private Sub Command3_Click()
    frmCalculations.Hide
    Call frmReport.Command1_Click
    frmReport.Show
End Sub
Private Sub Command4_Click()
  Dim Input1 As String
RestartThis:
    Input1 = InputBox("How many digits after the decimal place? Leave blank to quit", , Int(Precision.Caption))
    If LenB(Input1) = 0 Then
        '<:-):WARNING: Empty String comparision updated to use LenB()
        Exit Sub
    End If
    If IsNumeric(Int(Input1)) = False Then
        GoTo RestartThis
    End If
    Precision.Caption = Int(Input1)
    Call Command1_Click
    Call Text2_Change
    Call Text3_Change
    Call Text4_Change
End Sub
Private Sub Form_Load()
    Text1(0).Text = frmMain.Text1(0).Text
    Text1(1).Text = frmMain.Text1(1).Text
    Text1(2).Text = frmMain.Text1(3).Text
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    If IsNumeric(Text1(Index).Text) = False Then
        Select Case Index
         Case 0
            Text1(0).Text = 200
         Case 1
            Text1(1).Text = 9.8
         Case 2
            Text1(2).Text = 0
        End Select
    End If
End Sub
Private Sub Text2_Change()
    If IsNumeric(Text2.Text) = False Then
        Exit Sub
    End If
    Label2.Caption = Round(FindVelocity(Text1(1).Text, Text2.Text), Int(Precision.Caption)) & " m/s"
End Sub
Private Sub Text3_Change()
    If IsNumeric(Text3.Text) = False Then
        Exit Sub
    End If
    Label5.Caption = Round(FindInitialHeight(Text3.Text, Text1(2).Text, Text1(1).Text), Int(Precision.Caption))
End Sub
Private Sub Text4_Change()
    If IsNumeric(Text4.Text) = False Then
        Exit Sub
    End If
    Label6.Caption = Round(FindInitialVelocity(Text1(0).Text, Text4.Text, Text1(1).Text), Int(Precision.Caption))
End Sub

