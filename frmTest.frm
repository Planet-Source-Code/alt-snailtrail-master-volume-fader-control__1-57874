VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SnailTrail Volume Control"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   2115
      TabIndex        =   13
      Top             =   2685
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Demo:"
      Height          =   2610
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   2895
      Begin VB.TextBox txtSize 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   210
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "Use Gradient Colors"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Enable/Disable Gradient Colors"
         Top             =   1215
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.Timer tmrVis 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4185
         Top             =   585
      End
      Begin VB.PictureBox picStart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   1530
         Width           =   360
      End
      Begin VB.PictureBox picMid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for middle of the ColorBar"
         Top             =   1830
         Width           =   360
      End
      Begin VB.PictureBox picEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for top of the ColorBar"
         Top             =   2130
         Width           =   360
      End
      Begin VB.CheckBox chkSegment 
         Caption         =   "Use Segments"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Show solid ColorBar or in segments"
         Top             =   615
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin prjMasterVolume.MasterVolume MasterVolume1 
         Height          =   120
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   212
         BackColor       =   0
         Mute            =   -1  'True
         SliderIcon      =   "frmTest.frx":0000
         UseGradient     =   0   'False
      End
      Begin VB.PictureBox picFore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   1530
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblSize 
         Caption         =   "Segment Size (2 to 5)"
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label lblStart 
         Caption         =   "SnailTrail Start Gradient Color"
         Height          =   210
         Left            =   540
         TabIndex        =   8
         Top             =   1545
         Width           =   2100
      End
      Begin VB.Label lblMid 
         Caption         =   "SnailTrail Mid Gradient Color"
         Height          =   210
         Left            =   540
         TabIndex        =   7
         Top             =   1845
         Width           =   2040
      End
      Begin VB.Label lblEnd 
         Caption         =   "SnailTrail End Gradient Color"
         Height          =   210
         Left            =   540
         TabIndex        =   6
         Top             =   2145
         Width           =   2100
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4185
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGradient_Click()
     MasterVolume1.UseGradient = chkGradient.Value
     If chkGradient Then
          picFore.Visible = False
          lblStart = "SnailTrail Start Gradient Color"
          picStart.Visible = True
          picMid.Visible = True
          picEnd.Visible = True
          lblMid.Visible = True
          lblEnd.Visible = True
     Else
          picFore.Visible = True
          lblStart = "SnailTrail Color"
          picStart.Visible = False
          picMid.Visible = False
          picEnd.Visible = False
          lblMid.Visible = False
          lblEnd.Visible = False
     End If
End Sub

Private Sub chkSegment_Click()
     MasterVolume1.Segmented = chkSegment.Value
End Sub

Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     picStart.BackColor = MasterVolume1.GradientStartColor
     picMid.BackColor = MasterVolume1.GradientMidColor
     picEnd.BackColor = MasterVolume1.GradientEndColor
     picFore.BackColor = MasterVolume1.ForeColor
     MasterVolume1.Segmented = chkSegment
     MasterVolume1.UseGradient = chkGradient
     txtSize.Text = CStr(MasterVolume1.SegmentSize)
End Sub

Private Sub picEnd_Click()
     cdlg.CancelError = True
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     MasterVolume1.GradientEndColor = cdlg.Color
     picEnd.BackColor = cdlg.Color
End Sub

Private Sub picFore_Click()
     cdlg.CancelError = True
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     MasterVolume1.ForeColor = cdlg.Color
     picFore.BackColor = cdlg.Color
End Sub

Private Sub picMid_Click()
     cdlg.CancelError = True
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     MasterVolume1.GradientMidColor = cdlg.Color
     picMid.BackColor = cdlg.Color
End Sub

Private Sub picStart_Click()
     cdlg.CancelError = True
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     MasterVolume1.GradientStartColor = cdlg.Color
     picStart.BackColor = cdlg.Color
End Sub

Private Sub txtSize_Change()
     Select Case txtSize
          Case "2", "3", "4", "5"
               MasterVolume1.SegmentSize = CLng(txtSize)
          Case Else
               ' empty
               txtSize = ""
     End Select
End Sub
