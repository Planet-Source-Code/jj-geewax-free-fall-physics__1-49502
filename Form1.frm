VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Free Fall Physics - Graphical Version"
   ClientHeight    =   5835
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6915
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
   ScaleHeight     =   5835
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "More Stuff"
      Height          =   2895
      Left            =   1680
      TabIndex        =   15
      Top             =   2880
      Width           =   5175
      Begin VB.Frame Frame6 
         Caption         =   "Problem Solving"
         Height          =   2535
         Left            =   3600
         TabIndex        =   33
         Top             =   240
         Width           =   1455
         Begin VB.CommandButton Command4 
            Caption         =   "More"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Calculate!"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Calculate!"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "What is the velocity at t = ?"
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "When will it hit the ground?"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Points"
         Height          =   2535
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   1575
         Begin VB.ListBox HTX 
            Height          =   2205
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Mouse Coordinate"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         Begin MSComDlg.CommonDialog CMD1 
            Left            =   1560
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DefaultExt      =   "txt"
            Filter          =   "*.txt | Text Files"
         End
         Begin VB.Label MouseX 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label MouseY 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   720
            TabIndex        =   21
            Top             =   240
            Width           =   90
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Configurations"
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   840
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
            Index           =   0
            Left            =   720
            TabIndex        =   28
            Text            =   "200"
            Top             =   1320
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
            TabIndex        =   25
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
            Index           =   2
            Left            =   720
            TabIndex        =   22
            Text            =   "6"
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
            Index           =   3
            Left            =   720
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Height the object is dropped from."
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "m"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   29
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "m/s²"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   27
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Acc."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Acceleration Due to Gravity  (positive in this case!!!)"
            Top             =   960
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   24
            Top             =   600
            Width           =   45
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Iter."
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Number of Iterations"
            Top             =   600
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "m/s"
            Height          =   195
            Index           =   8
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Init Vel"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Initial Velocity"
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drop!"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ideas Employed in this Program"
      Height          =   2775
      Left            =   1680
      TabIndex        =   10
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text12 
         Height          =   2415
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   2895
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
         Left            =   120
         Picture         =   "Form1.frx":0032
         ScaleHeight     =   2340
         ScaleWidth      =   1905
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      DrawWidth       =   7
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MouseIcon       =   "Form1.frx":E8F4
      MousePointer    =   2  'Cross
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   105.206
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.Label lblT 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "lbT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line10 
         Index           =   2
         X1              =   69.414
         X2              =   69.414
         Y1              =   0
         Y2              =   199.604
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   270
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   30
         TabIndex        =   8
         Top             =   4800
         Width           =   90
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   7
         Top             =   2325
         Width           =   270
      End
      Begin VB.Line Line1 
         DrawMode        =   5  'Not Copy Pen
         X1              =   0
         X2              =   500
         Y1              =   98.833
         Y2              =   98.833
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   500
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   500
         Y1              =   199.604
         Y2              =   199.604
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   500
         Y1              =   49.416
         Y2              =   49.416
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   500
         Y1              =   148.21
         Y2              =   148.21
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   500
         Y1              =   24.688
         Y2              =   24.688
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   500
         Y1              =   74.105
         Y2              =   74.105
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   500
         Y1              =   123.521
         Y2              =   123.521
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   500
         Y1              =   172.938
         Y2              =   172.938
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   6
         Top             =   1050
         Width           =   270
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   30
         TabIndex        =   5
         Top             =   3570
         Width           =   180
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "75"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   4
         Top             =   2925
         Width           =   180
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "175"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   3
         Top             =   450
         Width           =   270
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "125"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   2
         Top             =   1680
         Width           =   270
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   30
         TabIndex        =   1
         Top             =   4170
         Width           =   180
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   32.538
         X2              =   32.538
         Y1              =   0
         Y2              =   199.604
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   103.977
         X2              =   103.977
         Y1              =   0
         Y2              =   199.604
      End
      Begin VB.Line Line10 
         Index           =   10
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   199.604
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export Points"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import Points"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCalc 
      Caption         =   "&Calculations"
      Begin VB.Menu mnuCalcGround 
         Caption         =   "Hit &Ground"
      End
      Begin VB.Menu mnuCalcInstVel 
         Caption         =   "Instant &Vel."
      End
      Begin VB.Menu mnuCalcMore 
         Caption         =   "&More Calculations"
      End
      Begin VB.Menu mnuCalcReport 
         Caption         =   "&Report"
         Begin VB.Menu mnuCalcReportShow 
            Caption         =   "&Show Report"
         End
         Begin VB.Menu mnuCalcReportPData 
            Caption         =   "&Print Data Points"
         End
         Begin VB.Menu mnuCalcReportSave 
            Caption         =   "Sa&ve Report"
         End
      End
   End
   Begin VB.Menu mnuExact 
      Caption         =   "More &Exact"
      Begin VB.Menu mnuExactExact 
         Caption         =   "&Exact Version"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpAir 
         Caption         =   "Air &Resistance?"
      End
      Begin VB.Menu mnuHelpWhat 
         Caption         =   "&What is the point?"
      End
      Begin VB.Menu mnuFileSAS 
         Caption         =   "&SAS"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'so we must dim everything

Private Sub Command1_Click()

  Dim T  As Integer
  Dim i  As Integer
  Dim UB As Integer
  Dim p  As Double

    '<:-):UPDATED: Multiple Dim line separated
    On Error Resume Next
    '<:-):RISK:  Load/UnLoad are safer with Error Trapping
    'this is where most of the code is, everything else is pretty pointless
    ' most of that is to make stuff look pretty and everything
    'scaleable axis was kinda fun to do, but a little bit of a pain...
    'comments go to jgeewax@standrews-de.org
    'dim our variables
    Picture1.Cls ' clear our picture box
    Call HTX.Clear ' clear our point list
    If lblT.Count > 1 Then ' if we have more than one label
        For i = 1 To lblT.UBound 'get all them
            Unload lblT(i) 'and unload them!!!
        Next i
    End If
    UB = Int(Val(Text1(2).Text)) 'set our upper bound to the # of
    'iterations inputted.
    'make sure its an integer
    For T = 0 To UB 'for 0 to the upper bound
        p = (Val(Text1(0).Text) - (Val(Text1(3).Text) * T + (0.5 * Val(Text1(1).Text) * T * T)))
        'this sets our p value with the equation to the left of the
        'about box
        'if we have a value that is really
        If p < -100 Then
            Exit Sub
        End If
        'far down, then simply stop, we dont want insignificant data
        Picture1.PSet (Picture1.ScaleWidth / 2, Picture1.ScaleHeight - p), RGB(0, 0, 0)
        'set our data to the picture in bright blue
        HTX.AddItem (p) 'add the y coordinate (height) to the HTX list box
        i = T + 1 ' set our value of i to t + 1 (so we dont get 0 in there..)
        If lblT.Count <= UB + 1 Then ' we are already looping, so
            'if we dont have all of the labels generated yet,
            'make some new ones
            Load lblT(i)
            lblT(i).Visible = True 'make sure theyre visible
        End If
        If i Mod 2 = 1 Then 'if the index is odd, put it to one side
            lblT(i).Left = (Picture1.ScaleWidth / 2) + 10
         ElseIf i Mod 2 = 0 Then 'if its even, put it to the other
            lblT(i).Left = (Picture1.ScaleWidth / 2) - 50
        End If
        lblT(i).Top = Picture1.ScaleHeight - p ' set the position
        '(height) of the label so it's next to the point
        lblT(i).Caption = Round(p, 1) 'set the caption of the label
        'to that points height coordinate
    Next T
    On Error GoTo 0
    '<:-):RISK: Turns off 'On Error Resume Next' in routine( Good coding but may not be what you want)

End Sub

Private Sub Command2_Click()

    Call frmCalculations.Command1_Click
    MsgBox "The object will hit the ground at time t = " & frmCalculations.Label1.Caption

End Sub

Private Sub Command3_Click()

  Dim X As String

m:
    X = InputBox("At what time do you want to find the velocity? Leave blank to exit.", , "2")
    If LenB(X) = 0 Then
        '<:-):WARNING: Empty String comparision updated to use LenB()
        Exit Sub
    End If
    If IsNumeric(X) = False Then
        GoTo m
    End If
    MsgBox ("The object is traveling at a velocity of " & Round(FindVelocity(Text1(1).Text, Val(X)), 1) & " m/s at time " & Round(X, 1) & " seconds.")

End Sub

Private Sub Command4_Click()
frmCalculations.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
  'This just makes the about box pretty and full of information

  Dim S1 As String
  Dim S2 As String
  Dim S3 As String
  Dim S4 As String

    '<:-):UPDATED: Multiple Dim line separated
    S1 = "For this program, we will employ the simple equations of 1 dimensional kinematics to have the computer predict the motion of a ""free falling"" object due to the force of the earth's gravity. We use the equations at right which can be easily derived using common integration techniques and some simple thinking about what acceleration, velocity, and displacement actually are."
    S2 = "Please feel free to email me with any comments at: " & vbNewLine & "jgeewax@standrews-de.org"
    S3 = vbNewLine & "More elaborate info:" & vbNewLine & vbNewLine & "The first equation is simply the standard acceleration due to gravity which has been measured to be 9.8 m/s²." & vbNewLine & "The second equation is used to determine the velocity at any one point." & vbNewLine & "The third equation determines the position or height from the ground at any point when starting from a point hi, with an acceleration due to gravity."
    S4 = "The last equation is the most complicated, which tells us to get the height from the ground like the others, but this one accounts for ann inital velocity vi that the object is being launched with. This one acts as if the object is thrown down toward the ground."
    Text12.Text = S1 & vbNewLine & S2 & vbNewLine & S3 & vbNewLine & S4
    Text12.Locked = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub mnuCalcGround_Click()
Call Command2_Click
End Sub

Private Sub mnuCalcInstVel_Click()
Call Command3_Click
End Sub

Private Sub mnuCalcMore_Click()

    frmCalculations.Show vbModal

End Sub

Private Sub mnuCalcReportPData_Click()

  Dim i As Integer

    On Error GoTo ErrHandler
    frmReport.rtbReport.Text = vbNullString
    '<:-):WARNING: Empty String assignment updated to use vbNullString
    For i = 0 To HTX.ListCount - 1
        frmReport.rtbReport.Text = frmReport.rtbReport.Text & (i + 1) & "     " & HTX.List(i) & vbNewLine
    Next i
    frmReport.rtbReport.SelPrint (Printer.hDC)
    frmReport.rtbReport.Text = vbNullString
    '<:-):WARNING: Empty String assignment updated to use vbNullString

Exit Sub

ErrHandler:
    '<:-):WARNING: Unneeded Exit Sub

End Sub

Private Sub mnuCalcReportPrint_Click()

  '<:-)Auto-inserted With End...With Structure

    With frmReport
        Call .Command1_Click
        .Show
        frmReport.PrintForm
    End With 'frmReport

End Sub

Private Sub mnuCalcReportSave_Click()
CMD1.CancelError = True
CMD1.ShowSave


    Call frmReport.Command1_Click
    frmReport.Show

Call SaveText(frmReport.rtbReport, CMD1.FileName)
End Sub

Private Sub mnuCalcReportShow_Click()

    Call frmReport.Command1_Click
    frmReport.Show

End Sub

Private Sub mnuExactExact_Click()
frmExact.Show
frmMain.Hide
End Sub

Private Sub mnuFileExit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo) = vbYes Then
    End
End If
End Sub

Private Sub mnuFileSAS_Click()
MsgBox "SAS is St. Andrew's School, and as you probably guessed, I attend there. This program was made for a class (AP Physics) that I am taking there currently."

End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "This program was made around 10/28/03 by JJ Geewax for AP Physics." & vbNewLine & _
"Please send all comments or suggestions to jgeewax@standrews-de.org"

End Sub

Private Sub mnuHelpAir_Click()
MsgBox "For this program, I neglected air resistance although it does apply because the definition of air resistance says that it is the force of air that acts in the opposite direction of the motion of the object." & vbNewLine & _
"In this case, the air resistance is just like the force of friction, which would in a ""free fall"" environment act in the +y or complete upward position." & vbNewLine & _
"This should make the air resistance easy to calculate, so I will do it sometime later, but for now Air Resistance is being left out of the equation completely!"

End Sub

Private Sub mnuHelpWhat_Click()
MsgBox "There really isn't a huge point to this program. I have been working on a number of Computer Applications to Physics projects and I have found that this one is idea to a beginner programmer and beginner Physicist." & vbNewLine & _
"The main point of this is to illustrate how the computer can be utilized to display graphically the 1-Dimensional Kinematics problems that back when they were still being discovered would have taken the physicist quite some time to calculate." & vbNewLine & _
"I hope that this can emphasize the extreme power of the computer as a tool for physics calculations. There will be many more programs like this to come. Thanks for using!"

End Sub

Private Sub Picture1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  'this tells the labels to display the appropriate coordinates
  'since this in 1-d motion, we only need the y, so i hid the
  'x coordinate label in runtime.

    MouseX.Caption = Round(X, 1)
    MouseY.Caption = Round(Picture1.ScaleHeight - Y, 1)

End Sub

Private Sub Text1_Change(Index As Integer)

  Dim Dummy As Double

    'whenever the text boxes change.
    If Index = 0 Then 'this is if the height is changed. to change the scaling.
        'this was kind of a pain, but still fun.
        'changed the height dropped from
        ' if its "" then get out
        If LenB(Text1(Index).Text) = 0 Then
            '<:-):WARNING: Empty String comparision updated to use LenB()
            Exit Sub
        End If
        'if its negative then get out
        If Text1(Index).Text <= 0 Then
            Exit Sub
        End If
        ' if it's less than 8 meters, then its too short
        If Text1(Index).Text < 8 Then
            Exit Sub
        End If
        'you could change this if you wanted.
        lblVal(0).Caption = Text1(Index).Text 'have the top label say the top height
        lblVal(8).Caption = "0" ' have the bottom label still say zero
        Dummy = Val(Text1(Index).Text / 8) ' make our division so we know what to subtract
        'this next part is mostly just a subtracting the dummy from the one above to spread the values out
        lblVal(1).Caption = Val(Round(lblVal(0).Caption - Dummy, 0))
        lblVal(2).Caption = Val(Round(lblVal(1).Caption - Dummy, 0))
        lblVal(3).Caption = Val(Round(lblVal(2).Caption - Dummy, 0))
        lblVal(4).Caption = Val(Round(lblVal(3).Caption - Dummy, 0))
        lblVal(5).Caption = Val(Round(lblVal(4).Caption - Dummy, 0))
        lblVal(6).Caption = Val(Round(lblVal(5).Caption - Dummy, 0))
        lblVal(7).Caption = Val(Round(lblVal(6).Caption - Dummy, 0))
        'make sure to round all of our captions, we dont want crazy long decimals!!!
        'change our scale height so everything is good
        'i love this property of the picture box!!!
        Picture1.ScaleHeight = Val(Round(Text1(Index).Text, 1)) 'round it so its not crazy decimalled
    End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)

    If Index = 0 Then 'if they click out of the height box (they are done editing)
        'do some simple things.
        'if its nothing, then set it to the default 200
        If LenB(Text1(Index).Text) = 0 Then
            '<:-):WARNING: Empty String comparision updated to use LenB()
            Text1(Index).Text = "200"
        End If
        'if its less than 8, set to the default 200
        If Text1(Index).Text < 8 Then
            Text1(Index).Text = "200"
        End If
    End If

End Sub
