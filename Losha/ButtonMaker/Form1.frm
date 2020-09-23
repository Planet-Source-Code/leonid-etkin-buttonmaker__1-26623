VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Button Maker"
   ClientHeight    =   5520
   ClientLeft      =   2265
   ClientTop       =   1575
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Lucida Blackletter"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   97.367
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   98.161
   Begin MSComDlg.CommonDialog Common1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox sizeTool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   840
      ScaleHeight     =   2160
      ScaleWidth      =   4125
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   4185
      Begin VB.TextBox txtWidth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Text            =   "21"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtHeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Text            =   "8"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblheight 
         Caption         =   "Height:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblSize 
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.PictureBox ActionTool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   840
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton optActNone 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Changes Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   40
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Shrinks and dissapears on first click"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   39
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Changes Size constantly (on every click)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   38
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Changes Size Once"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Makes Text Appear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Changes Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Disappears"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Click-Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox textTool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   840
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox ChkStrike 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "OCR-B-10 BT"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Strike Out"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox ChkUnder 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Miriam Transparent"
            Size            =   14.25
            Charset         =   177
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Underlined"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CheckBox ChkItalic 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "OCR-B-10 BT"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Italic"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Bold"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbSize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1080
         List            =   "Form1.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   915
         Width           =   1815
      End
      Begin VB.ComboBox cmbFont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0004
         Left            =   1080
         List            =   "Form1.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1080
         TabIndex        =   10
         Text            =   "Button1"
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Font:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Text:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Timer tmrShrink 
      Interval        =   1
      Left            =   240
      Top             =   2280
   End
   Begin VB.PictureBox colorTool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      ScaleHeight     =   2835
      ScaleWidth      =   4155
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   12
         Left            =   1920
         MaskColor       =   &H00008080&
         TabIndex        =   31
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   11
         Left            =   1920
         MaskColor       =   &H00008080&
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   10
         Left            =   1920
         MaskColor       =   &H00008080&
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   9
         Left            =   1080
         MaskColor       =   &H00008080&
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   8
         Left            =   1080
         MaskColor       =   &H00008080&
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   7
         Left            =   1080
         MaskColor       =   &H00008080&
         TabIndex        =   26
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   6
         Left            =   1080
         MaskColor       =   &H0000FF00&
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         MaskColor       =   &H0000FFFF&
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   240
         MaskColor       =   &H00C00000&
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   3
         Left            =   240
         MaskColor       =   &H000080FF&
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         MaskColor       =   &H00C00000&
         TabIndex        =   21
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   240
         MaskColor       =   &H00FF0000&
         TabIndex        =   20
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton OptColor 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         MaskColor       =   &H000000FF&
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton optNone 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Choose a color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox PositionTool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label6 
         Caption         =   "Choose a position for the button"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Next "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Button1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1191
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make a button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   1755
      Width           =   2655
   End
   Begin VB.Label lblTxtAppear 
      BackStyle       =   0  'Transparent
      Caption         =   "Appearing Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1800
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image imgFlash 
      Height          =   4485
      Left            =   720
      Picture         =   "Form1.frx":0008
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OKS As Integer
Dim iAction As Integer
Dim btmove As Boolean
Dim iActText As String
Dim iActTextF As String
Dim Xsize As Integer, Ysize As Integer
Dim iSizeChangeW As Integer, iSizeChangeH As Integer
Dim iConstantChangeW As Integer, iConstantChangeH As Integer
Dim IncDec As Integer


Private Sub Check2_Click()

End Sub

Private Sub chkBold_Click()
Command1.FontBold = Not Command1.FontBold
End Sub

Private Sub ChkItalic_Click()
Command1.FontItalic = Not Command1.FontItalic
End Sub

Private Sub ChkStrike_Click()
Command1.FontStrikethru = Not Command1.FontStrikethru
End Sub

Private Sub ChkUnder_Click()
Command1.FontUnderline = Not Command1.FontUnderline

End Sub

Private Sub cmbFont_Click()
On Error Resume Next
Command1.FontName = cmbFont.Text
End Sub

Private Sub cmbSize_Click()
Command1.FontSize = cmbSize.Text
End Sub

Private Sub cmdBack_Click()
If OKS = 4 Then
    cmdOk.Caption = "Next"
    cmdBack.Top = 33.867
    cmdOk.Top = 33.867
    cmdOk.Left = 52.917
    cmdOk.Visible = True
    Command1.Top = 2.117
    Command1.Left = 40.217
    lblTxtAppear.Visible = False
    ActionTool.Visible = True
    PositionTool.Visible = False
    btmove = False
    OKS = 3
ElseIf OKS = 3 Then
    colorTool.Visible = True
    ActionTool.Visible = False
    OKS = 2
ElseIf OKS = 2 Then
    sizeTool.Visible = True
    colorTool.Visible = False
    OKS = 1
ElseIf OKS = 1 Then
    textTool.Visible = True
    sizeTool.Visible = False
    cmdBack.Enabled = False
    OKS = 0
End If
End Sub

Private Sub cmdMake_Click()
imgFlash.Visible = False
Command1.Visible = True
textTool.Visible = True
cmdMake.Visible = False
cmdOk.Visible = True
End Sub

Private Sub cmdOk_Click()
If OKS = 0 Then
    cmdBack.Enabled = True
    Command1.Caption = txtText.Text
    textTool.Visible = False
    sizeTool.Visible = True
    OKS = 1
ElseIf OKS = 1 Then
    sizeTool.Visible = False
    colorTool.Visible = True
    OKS = 2
ElseIf OKS = 2 Then
    colorTool.Visible = False
    ActionTool.Visible = True
    OKS = 3
ElseIf OKS = 3 Then
    lblTxtAppear.Visible = False
    ActionTool.Visible = False
    PositionTool.Visible = True
    cmdOk.Caption = "Finish"
    cmdOk.Top = 74
    cmdBack.Top = 74
    btmove = True
    OKS = 4
Else
    cmdOk.Visible = False
    PositionTool.Visible = False
    cmdBack.Visible = False
    OKS = 5
End If

If OKS >= 1 And OKS < 5 Then
    cmdBack.Visible = True
End If
End Sub

Private Sub Command1_Click()

If btmove = False And _
PositionTool.Visible = True Then
    btmove = True
ElseIf PositionTool.Visible = False And _
OKS = 5 And cmdOk.Visible = False Then
Action_Execute
End If

End Sub

Private Sub Form_Click()
If btmove = True Then
    btmove = False
End If
If OKS = 3 And btmove = False And lblTxtAppear.Visible = True And cmdOk.Visible = False Then
    cmdOk.Visible = True
    ActionTool.Visible = True
End If
End Sub

Private Sub Form_Load()

Dim X, iItem
iItem = 6
For X = 0 To 3
    iItem = iItem + 2
    cmbSize.AddItem iItem
Next X

cmbSize.AddItem "18"
cmbSize.AddItem "24"

cmbFont.AddItem "Comic Sans MS"
cmbFont.AddItem "Courier New"
cmbFont.AddItem "MS Sans Serif"
cmbFont.AddItem "ShelleyAllegro BT"
cmbFont.AddItem "Old English"
cmbFont.AddItem "LcdD"
cmbFont.AddItem "Lucida Blackletter"
cmbFont.AddItem "Bloody"
cmbFont.AddItem "PosterBodoni BT"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If btmove = True Then
    Command1.Left = X + 3.2
    Command1.Top = Y + 3.2
ElseIf OKS = 3 And btmove = False And lblTxtAppear.Visible = True And _
cmdOk.Visible = False Then
    lblTxtAppear.Left = X + 3.2
    lblTxtAppear.Top = Y + 3.2
End If
End Sub

Private Sub lblTxtAppear_Click()
cmdOk.Visible = False
ActionTool.Visible = False
End Sub

Private Sub optAction_Click(Index As Integer)
On Error GoTo err_handler

iAction = optAction(Index).Index
If optAction(1).Value = True Then
   iActText = InputBox("Changes text to:")
ElseIf optAction(2).Value = True Then
   iActTextF = InputBox("Appearing Text:")
   If iActTextF <> "" Then
   X = MsgBox("Are You Sure?", vbYesNo + vbQuestion, "Action")
   If X = 6 Then
    lblTxtAppear.Visible = True
    ActionTool.Visible = False
    cmdOk.Visible = False
End If
End If
End If
If Index <> 2 Then
    lblTxtAppear.Visible = False
End If
If Index = 3 Then
    iSizeChangeW = InputBox("Width changes to:(In Millimiters)", "Width")
    iSizeChangeH = InputBox("Height changes to:(In Millimiters)", "Height")
End If
If Index = 4 Then
    iConstantChangeW = InputBox("Width constantly changes by:(In Millimiters)", "Width")
    iConstantChangeH = InputBox("Height constantly changes by:(In Millimiters)", "Height")
    IncDec = MsgBox("Increase or decrease in size? (Yes for increase, No for decrease)", vbYesNo)
End If
If Index = 6 Then
    Common1.ShowColor
End If
err_handler:
End Sub

Private Sub optActNone_Click()
iAction = -1
End Sub

Private Sub OptColor_Click(Index As Integer)
    Command1.BackColor = OptColor(Index).BackColor
End Sub

Private Sub optNone_Click()
    Command1.BackColor = &H8000000F
End Sub

Private Sub Timer1_Timer()

End Sub



Private Sub txtHeight_Change()
On Error GoTo err_Handler2
If txtHeight.Text > 50 Then
    MsgBox "The height of the button cannot exceed 50", vbInformation
    txtHeight.Text = 21
End If
Command1.Height = txtHeight.Text

err_Handler2:
End Sub

Private Sub txtText_Change()
Command1.Caption = txtText.Text
End Sub

Private Sub txtWidth_Change()
On Error GoTo err_handler
If txtWidth.Text > 50 Then
    MsgBox "The width of the button cannot exceed 50", vbInformation
    txtWidth.Text = 21
End If
Command1.Width = txtWidth.Text

err_handler:
End Sub


Public Sub Action_Execute()
On Error GoTo err_handler:
If iAction = 0 Then
    Command1.Visible = False
ElseIf iAction = 1 Then
    Command1.Caption = iActText
ElseIf iAction = 2 Then
    lblTxtAppear.Caption = iActTextF
    lblTxtAppear.Visible = True
ElseIf iAction = 3 Then
    Command1.Width = iSizeChangeW
    Command1.Height = iSizeChangeH
ElseIf iAction = 4 Then
If IncDec = 7 Then
    Command1.Height = Command1.Height - iConstantChangeH
    Command1.Width = Command1.Width - iConstantChangeW
Else
    Command1.Height = Command1.Height + iConstantChangeH
    Command1.Width = Command1.Width + iConstantChangeW
End If
ElseIf iAction = 5 Then
While Command1.Width > 5
    Command1.Move Command1.Left, Command1.Top, _
    Command1.Width - 1, Command1.Height - 1
    DoEvents
If Command1.Width <= 10 Then
    Command1.Visible = False
    Exit Sub
End If
Wend
ElseIf iAction = 6 Then
    Command1.BackColor = Common1.Color
End If
err_handler:
End Sub
