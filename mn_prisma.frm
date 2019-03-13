VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5AB778B2-4E89-4DCF-83B2-442F02E88CE6}#1.0#0"; "pngviewer.ocx"
Begin VB.Form prismasegitiga 
   Caption         =   "Prisma"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox sm 
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
      TabIndex        =   24
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox tP2 
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
      TabIndex        =   22
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox tP 
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
      TabIndex        =   21
      Top             =   4800
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "mn_prisma.frx":0000
      Top             =   360
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
      Left            =   7080
      TabIndex        =   20
      Top             =   6600
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
      Left            =   7080
      TabIndex        =   19
      Top             =   3600
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
      Left            =   7080
      TabIndex        =   18
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Luas Permukaan Prisma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   5400
      TabIndex        =   11
      Top             =   2640
      Width           =   5055
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BAD5
         TabIndex        =   25
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton hLP 
         Caption         =   "Hitung"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BB49
         TabIndex        =   14
         Top             =   3960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BBB1
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BC29
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BCA5
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BD1D
         TabIndex        =   23
         Top             =   360
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
      Top             =   6000
      Width           =   2535
   End
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
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume Prisma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   5055
      Begin VB.TextBox Text1 
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
         OleObjectBlob   =   "mn_prisma.frx":1BE69
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BED1
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BF49
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1BFC5
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "mn_prisma.frx":1C03D
         TabIndex        =   7
         Top             =   480
         Width           =   4575
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   2295
      Left            =   4560
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   4048
      Frame           =   4100
      Effects         =   "mn_prisma.frx":1C119
      BkgImage        =   "mn_prisma.frx":1C131
   End
End
Attribute VB_Name = "prismasegitiga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Prisma As New rumus

Private Sub Form_Load()
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub hLP_Click()
n2 = Prisma.LPPrisma(a2, tS2, tP2, sm)
End Sub

Private Sub hV_Click()
n = Prisma.VolPrisma(a, tS, tP)
End Sub
