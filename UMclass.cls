VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UMclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SI_metrico As Boolean

Function Conversione(ByVal inValore As Double, ByVal unita_in As UM, ByVal unita_out As UM, Optional ByVal Decimali As Byte) As Double
   Dim conv_in_temp As Double
   Dim conv_temp_out As Double
   
   conv_in_temp = 0
   conv_temp_out = 0
   Conversione = 0
   
   'selezione SI
   
   If SI_metrico Then
      If unita_in = UM.inch Then unita_in = UM.mm
      If unita_in = UM.lB Then unita_in = UM.kg
      If unita_in = UM.ft_min Then unita_in = UM.m_min
      
      If unita_out = UM.inch Then unita_out = UM.mm
      If unita_out = UM.lB Then unita_out = UM.kg
      If unita_out = UM.ft_min Then unita_out = UM.m_min
   End If
   
   'conversione ingresso --> m,kg
   
   Select Case unita_in
   Case UM.mm
       conv_in_temp = Round(inValore * 0.001, 6)         'mm-->m
   Case UM.mt
       conv_in_temp = Round(inValore, 6)                 'm-->m
   Case UM.kg
       conv_in_temp = Round(inValore, 6)                 'kg-->kg
   Case UM.inch
       conv_in_temp = Round(inValore * 0.0254, 6)        'inch-->m
   Case UM.ft
       conv_in_temp = Round(inValore * 0.3048, 6)        'ft-->mt
   Case UM.lB
       conv_in_temp = Round(inValore * 0.4536, 6)        'lb-->kg
   Case UM.mInch
       conv_in_temp = Round(inValore * 0.0000254, 6)     'minch-->m
   Case UM.yd
       conv_temp_out = Round(conv_in_temp * 0.9144, 6)   'm-->yd
   Case UM.dmm
       conv_in_temp = Round(inValore * 0.0001, 6)        'dmm-->m
   Case UM.ft_min
       conv_in_temp = Round(inValore * 0.00508, 6)       'ft/min-->m/s
   Case UM.inch_s
       conv_in_temp = Round(inValore * 0.0254, 6)        'inch/s-->m/s
   Case UM.m_min
       conv_in_temp = Round(inValore / 60, 6)            'm/min-->m/s
   Case UM.m_s
       conv_in_temp = Round(inValore, 6)                 'm/s-->m/s
   Case UM.ft_s
       conv_in_temp = Round(inValore * 0.3048, 6)        'ft/s-->m/s
   Case UM.inch_min
       conv_in_temp = Round(inValore * 0.00042333333, 6) 'inch/min-->m/s
   Case UM.dm
       conv_in_temp = Round(inValore * 0.1, 6)           'dm-->m
   Case UM.dm3
       conv_in_temp = Round(inValore * 0.001, 6)         'dm3-->m3
   Case UM.n
       conv_in_temp = Round(inValore * 0.102, 6)         'N-->Kg
   Case UM.inch3
       conv_in_temp = Round(inValore * 0.000016387064, 6) 'inch3-->m3
   Case UM.kg_dm3
       conv_in_temp = Round(inValore, 6)                  'kg/dm3-->Kg/dm3
   Case UM.n_dm3
       conv_in_temp = Round(inValore * 0.102, 6)          'N/dm3-->Kg/dm3
   Case UM.n_inch3
       conv_in_temp = Round(inValore * 6224.42189, 6)     'N/inch3-->Kg/dm3
   End Select
   
   'conversione m,kg ---> uscita
   
   Select Case unita_out
   Case UM.mm
       conv_temp_out = Round(conv_in_temp * 1000, 6)         'm-->mm
   Case UM.mt
       conv_temp_out = Round(conv_in_temp, 6)                'm-->m
   Case UM.kg
       conv_temp_out = Round(conv_in_temp, 2)                'kg-->kg
   Case UM.inch
       conv_temp_out = Round(conv_in_temp * 39.37, 6)        'm-->inch
   Case UM.ft
       conv_temp_out = Round(conv_in_temp * 3.28, 6)         'm-->ft
   Case UM.lB
       conv_temp_out = Round(conv_in_temp * 2.2046, 6)       'kg-->lb
   Case UM.mInch
       conv_temp_out = Round(conv_in_temp * 39.37 * 1000, 6) 'm-->minch
   Case UM.yd
       conv_temp_out = Round(conv_in_temp * 1.0936, 6)       'm-->yd
   Case UM.dmm
       conv_temp_out = Round(conv_in_temp * 10000, 6)             'dmm-->m
   Case UM.ft_min
       conv_temp_out = Round(conv_in_temp * 196.8503937, 6)       'm/s-->ft/min
   Case UM.inch_min
       conv_temp_out = Round(conv_in_temp * 2362.2047244, 6)      'm/s-->inch/min
   Case UM.m_min
       conv_temp_out = Round(conv_in_temp * 60, 6)                'm/s-->m/min
   Case UM.m_s
       conv_temp_out = Round(conv_in_temp, 6)                     'm/s-->m/s
   Case UM.ft_s
       conv_temp_out = Round(conv_in_temp * 3.280839, 6)          'm/s-->ft/s
   Case UM.inch_s
       conv_temp_out = Round(conv_in_temp * 39.37, 6)             'm/s-->inch/s
   Case UM.dm
       conv_temp_out = Round(conv_in_temp * 10, 6)                'm-->dm
   Case UM.dm3
       conv_temp_out = Round(conv_in_temp * 1000, 6)              'm3-->dm3
   Case UM.n
       conv_temp_out = Round(conv_in_temp * 9.807, 6)             'kg-->N
   Case UM.inch3
       conv_temp_out = Round(conv_in_temp * 61023.3779, 6)        'm3-->inch3
   Case UM.kg_dm3
       conv_temp_out = Round(conv_in_temp, 6)                     'kg/dm3-->Kg/dm3
   Case UM.n_dm3
       conv_temp_out = Round(conv_in_temp * 9.807, 6)             'kg/dm3-->N/dm3
   Case UM.n_inch3
       conv_temp_out = Round(conv_in_temp * 0.159144148, 6)       'kg/inch3-->N/dm3
   End Select
   
   Conversione = Round(conv_temp_out, CInt(Decimali))
   
End Function



