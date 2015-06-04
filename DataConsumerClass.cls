VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataConsumerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' DA FARE: dichiarare una variabile oggetto di tipo Recordset come fonte dati della
' classe. Ad esempio:

Private WithEvents RS As ADODB.Recordset
Attribute RS.VB_VarHelpID = -1

Public Property Get DataSource() As Recordset
    ' DA FARE: restituire l'oggetto fonte dati (DataSource) per la classe. Ad esempio:
   
    Set DataSource = RS
End Property
Public Property Set DataSource(ByVal objDataSource As Recordset)
    '  DA FARE: impostare l'oggetto fonte dati (DataSource) per la classe. Ad esempio:
    
    Set RS = objDataSource
    PropertyChanged "DataSource"
End Property
'Public Property Get DataMember() As DataMember
'    ' DA FARE: restituire il membro di dati (DataMember). Ad esempio:
'
'    Set DataMember = RS.DataMember
'End Property
'Public Property Let DataMember(ByVal DataMember As DataMember)
'    ' DA FARE: assegnare un nuovo membro di dati (DataMember). Ad esempio:
'
'    Set RS.DataMember = DataMember
'    PropertyChanged "DataMember"
'End Property

