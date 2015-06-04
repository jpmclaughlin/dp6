VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllarmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum AllarmiAttivi
    PagOrdini = 1
    PagStorico = 2
    PagLayout = 3
    PagPacco = 4
    PagRegge = 5
    PagPesa = 6
    Pagallarmi = 7
    PagSmusso = 8
    PagFiletto = 9
    PagWb = 10
    PagTaglio = 11
    PagEntrata = 12
    PagService = 13
End Enum

Public Enum AllType
   nessuno = 0
   allarmi = 1
   messaggi = 2
   entrambi = 3
End Enum

Public Tipo As AllType


