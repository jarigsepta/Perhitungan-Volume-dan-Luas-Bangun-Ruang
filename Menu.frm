VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Menu 
   Caption         =   "Rumus Bangun Ruang"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "Menu.frx":21488
      Top             =   240
   End
   Begin VB.Menu mn_file 
      Caption         =   "File"
      Begin VB.Menu mn_balok 
         Caption         =   "Balok"
      End
      Begin VB.Menu mn_tabung 
         Caption         =   "Tabung"
      End
      Begin VB.Menu mn_prisma 
         Caption         =   "Prisma"
      End
      Begin VB.Menu mn_limas 
         Caption         =   "Limas"
      End
   End
   Begin VB.Menu mn_keluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub mn_balok_Click()
    Me.Visible = False
    balok.Visible = True
End Sub

Private Sub mn_keluar_Click()
    x = MsgBox("Apakah Anda Ingin Keluar dari Program ?", vbQuestion + vbYesNo, "API2017B01")
    If x = vbYes Then
    End
    End If
End Sub

Private Sub mn_limas_Click()
    Me.Visible = False
    Limas.Visible = True
End Sub

Private Sub mn_prisma_Click()
    Me.Visible = False
    prismasegitiga.Visible = True
End Sub

Private Sub mn_tabung_Click()
    Me.Visible = False
    Tabung.Visible = True
End Sub

