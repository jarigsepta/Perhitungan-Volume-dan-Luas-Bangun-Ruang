VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5AB778B2-4E89-4DCF-83B2-442F02E88CE6}#1.0#0"; "pngviewer.ocx"
Begin VB.Form balok 
   Caption         =   "Balok"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10560
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
      Left            =   7080
      TabIndex        =   21
      Top             =   4920
      Width           =   2535
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
      TabIndex        =   20
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox l2 
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
      TabIndex        =   19
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Luas Permukaan Balok"
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
      Begin VB.TextBox p2 
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
         TabIndex        =   16
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton hLP 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   2760
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":0000
         TabIndex        =   12
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":0068
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":00D2
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":013A
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":01A6
         TabIndex        =   18
         Top             =   480
         Width           =   4815
      End
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
      Top             =   4920
      Width           =   2535
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
      TabIndex        =   6
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox l 
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
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume Balok"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":0286
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton hV 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2760
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":02EE
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":0358
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox p 
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
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":03C0
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_balok.frx":042C
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "mn_balok.frx":04CA
      Top             =   360
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   1335
      Left            =   3720
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      Frame           =   4100
      Effects         =   "mn_balok.frx":1BF9F
      BkgImage        =   "mn_balok.frx":1BFB7
   End
End
Attribute VB_Name = "balok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim balok As New rumus

Private Sub hLP_Click()
n2 = balok.LPBalok(p2, l2, t2)
End Sub

Private Sub hV_Click()
n = balok.VolBalok(p, l, t)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text4_Change()

End Sub

