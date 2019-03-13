VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5AB778B2-4E89-4DCF-83B2-442F02E88CE6}#1.0#0"; "pngviewer.ocx"
Begin VB.Form tabung 
   Caption         =   "Tabung"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "mn_tabung.frx":0000
      Top             =   360
   End
   Begin VB.TextBox t2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox r2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Luas Permukaan Tabung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5400
      TabIndex        =   11
      Top             =   1560
      Width           =   5055
      Begin VB.CommandButton hLP 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox phi2 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Text            =   "3.14"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox n2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   3360
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BAD5
         TabIndex        =   15
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BB3D
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BBA7
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BC17
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BC7B
         TabIndex        =   19
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox r 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume Tabung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox n 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox phi 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Text            =   "3.14"
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton hV 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   2760
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BD41
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BDA9
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BE13
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BE83
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_tabung.frx":1BEE7
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   1215
      Left            =   4680
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Frame           =   4100
      Effects         =   "mn_tabung.frx":1BF89
      BkgImage        =   "mn_tabung.frx":1BFA1
   End
End
Attribute VB_Name = "tabung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tabung As New rumus

Private Sub Form_Load()
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub hLP_Click()
n2 = Tabung.LPTabung(r2, t2)
End Sub

Private Sub hV_Click()
n = Tabung.VolTabung(r, t)
End Sub

