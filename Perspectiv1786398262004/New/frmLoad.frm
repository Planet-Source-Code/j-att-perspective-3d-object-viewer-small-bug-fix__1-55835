VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   34
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3600
      TabIndex        =   33
      Top             =   5040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1080
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraCanvas 
      Caption         =   "Canvas"
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   4695
      Begin VB.TextBox txtValue 
         Height          =   285
         Index           =   8
         Left            =   3000
         TabIndex        =   32
         Tag             =   "Canvas Width"
         Text            =   "300"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Index           =   7
         Left            =   720
         TabIndex        =   31
         Tag             =   "Canvas Height"
         Text            =   "300"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width:"
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraRendering 
      Caption         =   "Rendering"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4695
      Begin VB.Frame fraLight 
         Caption         =   "Light"
         Height          =   1695
         Left            =   2400
         TabIndex        =   12
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   19
            Tag             =   "Light Z"
            Text            =   "30"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   18
            Tag             =   "Light Y"
            Text            =   "0"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   17
            Tag             =   "Light X"
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chkLighted 
            Caption         =   "Lighted"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblLZ 
            Caption         =   "Z:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblLY 
            Caption         =   "Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblLX 
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame fraPosition 
         Caption         =   "Position"
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   27
            Tag             =   "Zoom"
            Text            =   "1 "
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   25
            Tag             =   "Z Position"
            Text            =   "-60"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   24
            Tag             =   "Y Position"
            Text            =   "0"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   23
            Tag             =   "X Position"
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblZoom 
            Caption         =   "Zoom:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblPZ 
            Caption         =   "Z:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lbPY 
            Caption         =   "Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblPX 
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.CheckBox chkZOrder 
         Caption         =   "Z Order"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmLoad.frx":08CA
         Left            =   120
         List            =   "frmLoad.frx":08DD
         TabIndex        =   9
         Text            =   "Style"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraObject 
      Caption         =   "3d Object"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblVersion 
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lbldVersion 
         Caption         =   "Version:"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblFilename 
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   240
      Picture         =   "frmLoad.frx":091C
      Top             =   4920
      Width           =   480
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Filename As String

Private Sub cmdCancel_Click()
    
    frmView.p_Loaded = False
    Unload Me
    
End Sub

Private Sub cmdLoad_Click()
    
    Dim i As Integer
    
    If Filename = "" Then
        MsgBox "Object file was not supplied or invalid!", vbOKOnly + vbCritical, "Error:"
        Exit Sub
    End If

    For i = 0 To 8
        If Not IsNumeric(txtValue(i).Text) Then
            MsgBox "The " & txtValue(i).Tag & " value you supplied is invalid!", vbOKOnly + vbCritical, "Error:"
            Exit Sub
        End If
    Next i
    
    If cmbStyle.ListIndex = -1 Then
        MsgBox "The style value was not supplied or is invalid!", vbOKOnly + vbCritical, "Error:"
        Exit Sub
    End If
    
    With frmView
        .p_Style = cmbStyle.ListIndex
        .p_ZOrder = chkZOrder.Value
        .p_Lighted = chkLighted.Value
        .p_Object = Filename
        .p_Zoom = txtValue(0).Text
        .p_coordX = txtValue(1).Text
        .p_coordY = txtValue(2).Text
        .p_coordZ = txtValue(3).Text
        .p_LightX = txtValue(4).Text
        .p_LightY = txtValue(5).Text
        .p_LightZ = txtValue(6).Text
        .p_CHeight = txtValue(7).Text
        .p_CWidth = txtValue(8).Text
        .p_Loaded = True
    End With
    
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    
    Dim strTemp As String
    
    With CD
        '.hWnd = Me.hWnd
        .DefaultExt = "odf"
        .Filter = "Object Definition File(*.ODF) | *.odf|All Files (*.*) | *.*"
        .ShowOpen
    End With
    Filename = CD.Filename
    
    Open Filename For Input As 1
        Input #1, strTemp
        If strTemp <> "3D OBJECT DEFINITION FILE" Then
            MsgBox "Not a valid object file!", vbOKOnly + vbCritical, "Open"
            Filename = ""
            Exit Sub
        End If
    
        'get version
        Input #1, strTemp
        lblVersion.Caption = Trim$(strTemp)
    
        'get name
        Input #1, strTemp
        lblName.Caption = Trim$(strTemp)
    Close #1
    
    txtFilename.Text = Filename
    
End Sub

