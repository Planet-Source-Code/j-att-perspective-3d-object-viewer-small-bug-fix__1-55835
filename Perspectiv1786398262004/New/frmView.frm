VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Viewer"
   ClientHeight    =   4995
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6780
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Render 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   4500
      Left            =   2160
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   90
      Width           =   4500
   End
   Begin VB.PictureBox SBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   452
      TabIndex        =   2
      Top             =   4680
      Width           =   6780
      Begin VB.Label lblT 
         Caption         =   "Translation:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblTrX 
         Height          =   255
         Left            =   960
         TabIndex        =   40
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblTrY 
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblTrZ 
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblR 
         Caption         =   "Rotation:"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblRotZ 
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblRotY 
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblRotX 
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   60
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   472
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox NBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   0
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   360
         TabIndex        =   31
         Text            =   "100"
         Top             =   4320
         Width           =   735
      End
      Begin VB.CommandButton cmdNextFrame 
         Caption         =   "Next Frame"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Timer tmrAuto 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   3600
      End
      Begin VB.TextBox txtTrZ 
         Height          =   285
         Left            =   360
         TabIndex        =   25
         Text            =   "-60"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtTrY 
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Text            =   "0"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtTrX 
         Height          =   285
         Left            =   360
         TabIndex        =   23
         Text            =   "0"
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox chkTrX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   2400
         Width           =   225
      End
      Begin VB.CheckBox chkTrY 
         Caption         =   "Check2"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   2760
         Width           =   225
      End
      Begin VB.CheckBox chkTrZ 
         Caption         =   "Check3"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   3120
         Width           =   225
      End
      Begin VB.TextBox txtTrXinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtTrYinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Text            =   "0"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtTrZinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Text            =   "0"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtZRot 
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtYRot 
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtXRot 
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkRotX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   225
      End
      Begin VB.CheckBox chkRotY 
         Caption         =   "Check2"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   225
      End
      Begin VB.CheckBox chkRotZ 
         Caption         =   "Check3"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1440
         Width           =   225
      End
      Begin VB.TextBox txtRotXinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtRotYinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtRotZinc 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblTranslation 
         Caption         =   "Translation:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Z:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblRotation 
         Caption         =   "Rotation:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblRZ 
         Caption         =   "Z:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblRY 
         Caption         =   "Y:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblRX 
         Caption         =   "X:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   136
         X2              =   136
         Y1              =   16
         Y2              =   328
      End
      Begin VB.Label lblNBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Negotiate Bar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuNegotiate 
         Caption         =   "Negotiate Bar"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status Bar"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private obj0            As New cls3dObject
Private Xstart          As Single
Private Ystart          As Single

Public p_Loaded          As Boolean
Public p_Style            As Integer
Public p_coordX           As Double
Public p_coordY           As Double
Public p_coordZ           As Double
Public p_Zoom             As Double
Public p_LightX           As Single
Public p_LightY           As Single
Public p_LightZ           As Single
Public p_Lighted          As Boolean
Public p_Object           As String
Public p_ZOrder           As Boolean
Public p_CWidth           As Long
Public p_CHeight          As Long

Private Sub chkAuto_Click()

    tmrAuto.Enabled = chkAuto.Value

End Sub

Private Sub cmdNextFrame_Click()
    
    Increment
    RenderPic
    Reset
    DoEvents

End Sub



Private Sub Form_Load()
    
    With NBar
        .Width = 1
        .Visible = False
    End With
    
    With SBar
        .Height = 1
        .Visible = False
    End With
    
End Sub

Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuContents_Click()
    frmContents.Show
End Sub

Private Sub mnuExit_Click()
    Unload frmLoad
    Unload frmAbout
    Unload Me
End Sub

Private Sub mnuLoad_Click()
    
    frmLoad.Show vbModal
    
    cmdNextFrame.Enabled = p_Loaded
    chkAuto.Enabled = p_Loaded
    
    If p_Loaded Then
        With Render
            .Height = p_CHeight
            .Width = p_CWidth
        End With
        NBar_Resize
        SBar_Resize
        LoadObj
    Else
        Render.Cls
    End If
    
End Sub

Private Sub mnuNegotiate_Click()
    
    With NBar
    
        If .Width = 1 Then
            .Width = 139
        Else
            .Width = 1
        End If
        .Visible = Not (.Visible)
        
    End With

End Sub

Private Sub mnuStatus_Click()
    
    With SBar
    
        If .Height = 1 Then
            .Height = 21
        Else
            .Height = 1
        End If
        .Visible = Not (.Visible)
        
    End With
    
End Sub

Private Sub NBar_Resize()
    
    lblNBar.Width = NBar.Width
    
    With Line2
        .X1 = lblNBar.Width - 1
        .X2 = lblNBar.Width - 1
        .Y2 = NBar.Height
    End With
    
    Render.Left = NBar.Width + 6
    
    Me.Width = (Render.Left + Render.Width + 12) * 15
    
End Sub

Private Sub Render_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Xstart = X
    Ystart = Y

End Sub

Private Sub Render_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If p_Loaded Then
        If Button <> 0 Then
            If Button = 1 Then
               txtYRot.Text = CSng(txtYRot.Text) + (X - Xstart)
               txtXRot.Text = CSng(txtXRot.Text) + (Y - Ystart)
            ElseIf Button = 2 Then 'NOT BUTTON...
               txtTrX.Text = CSng(txtTrX.Text) + ((X - Xstart) / 5) '- b/c of messed up x/y orientation of picbox
               txtTrY.Text = CSng(txtTrY.Text) + ((Y - Ystart) / 5) '/5 b/c it moves 2 fast otherwise
            End If
            Xstart = X
            Ystart = Y
            RenderPic
            Reset
        End If
    End If

End Sub

Private Sub SBar_Resize()

    Line1.X2 = SBar.Width
    
    Me.Height = (SBar.Height + Render.Height + 12) * 15 + 795

    
End Sub

Private Sub tmrAuto_Timer()

    cmdNextFrame_Click

End Sub

Private Sub Increment()

    If (NumCheck(txtRotXinc.Text, txtRotYinc.Text, txtRotZinc.Text, "Rotation Increments") = True) Or (NumCheck(txtTrXinc.Text, txtTrYinc.Text, txtTrZinc.Text, "Translation Increments") = True) Then
        Exit Sub
    End If
    If chkRotX.Value Then
        txtXRot.Text = CSng(txtRotXinc.Text) + CSng(txtXRot.Text)
    End If
    If chkRotY.Value Then
        txtYRot.Text = CSng(txtRotYinc.Text) + CSng(txtYRot.Text)
    End If
    If chkRotZ.Value Then
        txtZRot.Text = CSng(txtRotZinc.Text) + CSng(txtZRot.Text)
    End If
    If chkTrX.Value Then
        txtTrX.Text = CSng(txtTrXinc.Text) + CSng(txtTrX.Text)
    End If
    If chkTrY.Value Then
        txtTrY.Text = CSng(txtTrYinc.Text) + CSng(txtTrY.Text)
    End If
    If chkTrZ.Value Then
        txtTrZ.Text = CSng(txtTrZinc.Text) + CSng(txtTrZ.Text)
    End If

End Sub

Private Function NumCheck(ByVal strText1 As String, _
                          ByVal strText2 As String, _
                          ByVal strText3 As String, _
                          ByVal Description As String) As Boolean

    If Not (IsNumeric(strText1)) Or Not (IsNumeric(strText2)) Or Not (IsNumeric(strText3)) Then
        MsgBox Description & " can only be numeric values", vbOKOnly, "Error"
        NumCheck = True
        chkAuto.Value = vbUnchecked
    End If

End Function

Private Sub RenderPic()

    NumCheck txtXRot.Text, txtYRot.Text, txtZRot.Text, "Rotations"
    NumCheck txtTrX.Text, txtTrY.Text, txtTrZ.Text, "Translations"
    obj0.Rotate txtXRot.Text, txtYRot.Text, txtZRot.Text
    obj0.Translate txtTrX.Text, txtTrY.Text, txtTrZ.Text
    Render.Cls
    obj0.RenderObject

End Sub

Private Sub Reset()

    txtXRot.Text = obj0.RotateX
    txtYRot.Text = obj0.RotateY
    txtZRot.Text = obj0.RotateZ
    lblRotX.Caption = obj0.RotateX
    lblRotY.Caption = obj0.RotateY
    lblRotZ.Caption = obj0.RotateZ
    txtTrX.Text = obj0.TranslateX
    txtTrY.Text = obj0.TranslateY
    txtTrZ.Text = obj0.TranslateZ
    lblTrX.Caption = obj0.TranslateX
    lblTrY.Caption = obj0.TranslateY
    lblTrZ.Caption = obj0.TranslateZ

End Sub

Private Sub txtSpeed_Change()

    If IsNumeric(txtSpeed.Text) Then
        tmrAuto.Interval = txtSpeed.Text
    End If

End Sub

Public Sub LoadObj()

    obj0.LoadObject p_Object, Render, p_Style, p_coordX, p_coordY, p_coordZ, p_Zoom, 0, 0, 0, p_ZOrder, p_Lighted, p_LightX, p_LightY, p_LightZ

End Sub
