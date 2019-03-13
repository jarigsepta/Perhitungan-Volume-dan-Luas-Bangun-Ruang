VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rumus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function VolBalok(ByVal p As Integer, ByVal l As Integer, ByVal t As Integer)
VolBalok = p * l * t
End Function

Public Function LPBalok(ByVal p2 As Integer, ByVal l2 As Integer, ByVal t2 As Integer)
LPBalok = 2 * (p2 * l2) + (p2 * t2) + (l2 * t2)
End Function

Public Function VolTabung(ByVal r As Integer, ByVal t As Integer)
VolTabung = 3.14 * r ^ 2 * t
End Function

Public Function LPTabung(ByVal r2 As Integer, ByVal t As Integer)
LPTabung = 2 * 3.14 * r2 * (t2 + r2)
End Function

Public Function VolPrisma(ByVal a As Integer, ByVal tS As Integer, ByVal tP As Integer)
VolPrisma = (1 / 2 * a * tS) * tP
End Function

Public Function LPPrisma(ByVal a2 As Integer, ByVal tS2 As Integer, ByVal tP2 As Integer, ByVal sm As Integer)
LPPrisma = 2 * (1 / 2 * a2 * tS2) + ((a2 + tS2 + sm) + tP2)
End Function

Public Function VolLimas(ByVal a As Integer, ByVal tS As Integer, ByVal tL As Integer)
VolLimas = 1 / 3 * (1 / 2 * a * tS) * tL
End Function

Public Function LPLimas(ByVal a2 As Integer, ByVal tS2 As Integer, ByVal tL2 As Integer)
LPLimas = (1 / 2 * a2 * tS2) + 3 * (1 / 2 * a2 * tL2)
End Function
