Attribute VB_Name = "TrovaFiles"
Option Explicit
Const MAX_PATH = 260

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Const Invalid_Handle_Value = -1
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Type Filetrovato
    nomefile As String
    percorso As String
    dimensione As Long
End Type

Public arrayRicerca() As Filetrovato
Public contaTrovati As Double
Public AnnullaRicerca As Boolean

Public Function TagliaNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TagliaNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  TagliaNull = startstr
End Function


Public Sub ScovaTuttiFile(filenome As String, InizioDir As String)
    Dim hfile As Long
    Static SfileNome As String
    Static Sfnome As String
    Static f As WIN32_FIND_DATA
    Static fboolean As Boolean
    
    'trova il primo file
    contaTrovati = 0
    ReDim arrayRicerca(contaTrovati)
    hfile = FindFirstFile(InizioDir & "*.*", f)
    If hfile <> Invalid_Handle_Value Then fboolean = True
    Do
    DoEvents
    If AnnullaRicerca = True Then Exit Sub
    Sfnome = TagliaNull(f.cFileName)
    'per ora solo il nome di un file preciso
    'escludo cartelle di ogni genere
    If Not (f.dwFileAttributes = vbDirectory) And Not (f.dwFileAttributes = vbDirectory + vbHidden) And Not (f.dwFileAttributes = vbDirectory + vbArchive) And Not (f.dwFileAttributes = vbDirectory + vbSystem) And Not ((f.dwFileAttributes - vbDirectory) < (vbSystem + vbHidden + vbReadOnly)) Then
        'diversifico secondo il tipo di ricerca impostato
        'tutti i files: - *.*
        If filenome = "*.*" Then
            contaTrovati = contaTrovati + 1
           ' mTotRighe = contatrovati
            ReDim Preserve arrayRicerca(contaTrovati)
            arrayRicerca(contaTrovati).percorso = InizioDir
            arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
            arrayRicerca(contaTrovati).nomefile = Sfnome
        Else
    'tipo di ricerca:
    '0)  prova* - tutti i file che iniziano per prova ,
    '1)  *prova - tutti i files che finiscono per prova
    '2)  *prova* - tutti i files che contengono prova
    '3)  prova*prova - tutti i files che iniziano e finiscono con prova
        If InStr(1, filenome, "*") <> 0 Then
            Static pos1 As Integer
            Static pos2 As Integer
            Static str1 As String
            Static str2 As String
            Static lunghezzastr As Integer
            pos1 = 0
            pos2 = 0
            str1 = ""
            str2 = ""
            pos1 = InStr(1, filenome, "*")
            pos2 = InStr(pos1 + 1, filenome, "*")
    'tipo di ricerca 0
           If pos1 = Len(filenome) And pos2 = 0 Then
                str1 = Left(Trim(filenome), Len(Trim(filenome)) - 1)
                
                lunghezzastr = Len(filenome) - 1
                If Len(Sfnome) >= lunghezzastr Then
                    If LCase(Left(Trim(Sfnome), lunghezzastr)) = LCase(str1) Then
                    contaTrovati = contaTrovati + 1
             
                    ReDim Preserve arrayRicerca(contaTrovati)
                    arrayRicerca(contaTrovati).percorso = InizioDir
                    arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
                    arrayRicerca(contaTrovati).nomefile = Sfnome
                    End If
                End If
            Else
            
    'tipo di ricerca 1
            If pos1 = 1 And pos2 = 0 Then
                str1 = Right(Trim(filenome), Len(Trim(filenome)) - 1)
                lunghezzastr = Len(filenome) - 1
                If Len(Sfnome) >= lunghezzastr Then
                    If LCase(Right(Trim(Sfnome), lunghezzastr)) = LCase(str1) Then
                    contaTrovati = contaTrovati + 1
                  
                    ReDim Preserve arrayRicerca(contaTrovati)
                    arrayRicerca(contaTrovati).percorso = InizioDir
                    arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
                    arrayRicerca(contaTrovati).nomefile = Sfnome
                    End If
                End If
            Else
    'tipo di ricerca 2
            If pos1 = 1 And pos2 = Len(Trim(filenome)) Then
            str1 = Left(Right(filenome, Len(filenome) - 1), Len(filenome) - 2)
            If InStr(1, LCase(Sfnome), LCase(str1)) Then
                contaTrovati = contaTrovati + 1
              
                ReDim Preserve arrayRicerca(contaTrovati)
                arrayRicerca(contaTrovati).percorso = InizioDir
                arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
                arrayRicerca(contaTrovati).nomefile = Sfnome
            End If
            Else
    'tipo di ricerca 3
            
    'inizio
            str1 = LCase(Left(Trim(filenome), pos1 - 1))
    'fine
            str2 = LCase(Right(Trim(filenome), Len(Trim(filenome)) - pos1))
            If Len(Sfnome) >= Len(str1) + Len(str2) Then
            Debug.Print Sfnome
            Debug.Print LCase(Left(Sfnome, Len(str1))) & "  *  " & LCase(Right(Sfnome, Len(str2)))
            If LCase(Left(Sfnome, Len(str1))) = str1 And LCase(Right(Sfnome, Len(str2))) = str2 Then
                contaTrovati = contaTrovati + 1
    
                ReDim Preserve arrayRicerca(contaTrovati)
                arrayRicerca(contaTrovati).percorso = InizioDir
                arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
                arrayRicerca(contaTrovati).nomefile = Sfnome
            End If
        End If
        End If
        End If
        End If
    Else
    'tipo di ricerca di un nome univoco
                If LCase(Sfnome) = LCase(Trim(filenome)) Then
                contaTrovati = contaTrovati + 1

                ReDim Preserve arrayRicerca(contaTrovati)
                arrayRicerca(contaTrovati).percorso = InizioDir
                arrayRicerca(contaTrovati).dimensione = f.nFileSizeLow \ 1024 + 1
                arrayRicerca(contaTrovati).nomefile = Sfnome
                End If
    End If
    End If
    End If
    Loop While FindNextFile(hfile, f)
    fboolean = FindClose(hfile)
End Sub
