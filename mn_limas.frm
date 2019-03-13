VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5AB778B2-4E89-4DCF-83B2-442F02E88CE6}#1.0#0"; "pngviewer.ocx"
Begin VB.Form limas 
   Caption         =   "Limas"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7320
      TabIndex        =   21
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox tS2 
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
      Left            =   7320
      TabIndex        =   19
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Luas Permukaan Limas"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   1440
      Width           =   5175
      Begin VB.TextBox tL2 
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
         TabIndex        =   20
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton hLP 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox a2 
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
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":0000
         TabIndex        =   14
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":0068
         TabIndex        =   15
         Top             =   2160
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":00DE
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":015A
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":01C0
         TabIndex        =   18
         Top             =   480
         Width           =   4935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "mn_limas.frx":02AE
      Top             =   480
   End
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox tL 
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
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox tS 
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
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume Limas"
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
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox a 
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
         OleObjectBlob   =   "mn_limas.frx":1BD83
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":1BDEB
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":1BE61
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":1BEDD
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_limas.frx":1BF43
         TabIndex        =   7
         Top             =   480
         Width           =   4455
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   1215
      Left            =   3840
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2143
      Frame           =   4100
      Effects         =   "mn_limas.frx":1C017
      BkgImage        =   "mn_limas.frx":1C02F
   End
End
Attribute VB_Name = "limas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Limas As New rumus

Private Sub Form_Load()
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub hLP_Click()
n2 = Limas.LPLimas(a2, tS2, tL2)
End Sub

Private Sub hV_Click()
n = Limas.VolLimas(a, tS, tL)
End Sub
