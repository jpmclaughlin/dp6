VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrdineAttuale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'definizione ordine corrente nella pagina

Property Get CodEntrata() As Integer
   If Not (DB450 Is Nothing) Then CodEntrata = DB450.Word(0)
End Property
Property Get CodWB() As Integer
   If Not (DB460 Is Nothing) Then CodWB = DB460.Word(0)
End Property
Property Get CodPacco() As Integer
   If Not (DB470 Is Nothing) Then CodPacco = DB470.Word(0)
End Property
Property Get CodMPS() As Integer
   If Not (DB473 Is Nothing) Then CodMPS = DB473.Word(0)
End Property
Property Get CodRegge() As Integer
   If Not (DB480 Is Nothing) Then CodRegge = DB480.Word(0)
End Property
Property Get CodStoccaggio() As Integer
   If Not (DB486 Is Nothing) Then CodStoccaggio = DB486.Word(0)
End Property
Property Get Cod402() As Integer
   If Not (DB402 Is Nothing) Then Cod402 = DB402.Word(18)
End Property
Property Get CodLav() As Integer
   If Not (DB465 Is Nothing) Then CodLav = DB465.Word(0)
End Property


