Attribute VB_Name = "LibFunctions"
Option Compare Database

'---
Global Const dll_db_lib = "Purchase_dll_Corporate"
'---

 
'JAT 28-05-15. #Access2013
Function CurrentDbLib(Optional dblib) As Database
    
    
    On Error Resume Next
    
    Dim Path As String
    Dim fs, f, c
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set f = fs.GetFile(CurrentDb.Name)
      
    If IsMissing(dblib) Then
        dblib = LCase(Nz(dll_db_lib))
    Else
        dblib = LCase(Nz(dblib))
    End If
    
    If Left(f.Name, Len(dblib)) <> dblib Then
        Path = PATH_FERP_APP_LIB & dblib & "." & AppFileExtension
    Else
        Path = CurrentDb.Name
    End If
    Set CurrentDbLib = DBEngine.OpenDatabase(Path)
    
End Function

Public Function DSumLib(Expr As String, domain As String, Optional Criteria As String = "") As Double
    Dim db As Database
    Dim rs As Recordset
    Dim where As String
    Dim sql As String
    
On Error GoTo Function_Err
    Set db = CurrentDbLib(dll_db_lib)
    If Nz(Criteria, "") <> "" Then
        where = "WHERE " & Criteria
    End If
    
    sql = "Select sum(" & Expr & ") as results from " & domain & IIf(Nz(Criteria) <> "", " where " & Criteria, " ")
            
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        DSumLib = 0
    Else
        DSumLib = Nz(rs!results, 0)
    End If
    rs.Close
    db.Close
    
    Exit Function
Function_Err:
    Select Case err.Number
        Case 3265:
            Resume Next
        Case Else:
            fmsgbox "1287;i;c", err.Description
            Resume
    End Select
End Function

Public Function DCountLib(Expr As String, domain As String, Optional Criteria As String = "") As Double
    Dim db As Database
    Dim rs As Recordset
    Dim where As String
    Dim sql As String
On Error GoTo Function_Err
    Set db = CurrentDbLib(dll_db_lib)
    If Nz(Criteria, "") <> "" Then
        where = "WHERE " & Criteria
    End If
    
    sql = "Select count(" & Expr & ") as results from " & domain & IIf(Nz(Criteria) <> "", " where " & Criteria, " ")
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        DCountLib = 0
    Else
        DCountLib = rs!results
    End If
    
    rs.Close
    db.Close
    
    Exit Function
Function_Err:
    Select Case err.Number
        Case 3265:
            Resume Next
        Case Else:
    End Select
End Function

Public Function DMaxLib(Expr As String, domain As String, Optional Criteria As String = "") As Double
    Dim db As Database
    Dim rs As Recordset
    Dim where As String
    Dim sql As String
On Error GoTo Function_Err
    Set db = CurrentDbLib(dll_db_lib)
    If Nz(Criteria, "") <> "" Then
        where = "WHERE " & Criteria
    End If
    
    sql = "Select max(" & Expr & ") as results from " & domain & IIf(Nz(Criteria) <> "", " where " & Criteria, " ")
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        DMaxLib = 0
    Else
        DMaxLib = Nz(rs!results, 0)
    End If
    
    rs.Close
    db.Close
    
    Exit Function
Function_Err:
    Select Case err.Number
        Case 3265:
            Resume Next
        Case Else:
    End Select
End Function

Public Function DMinLib(Expr As String, domain As String, Optional Criteria As String = "") As Double
    Dim db As Database
    Dim rs As Recordset
    Dim where As String
    Dim sql As String
On Error GoTo Function_Err
    Set db = CurrentDbLib(dll_db_lib)
    If Nz(Criteria, "") <> "" Then
        where = "WHERE " & Criteria
    End If
    
    sql = "Select min(" & Expr & ") as results from " & domain & IIf(Nz(Criteria) <> "", " where " & Criteria, " ")
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        DMinLib = 0
    Else
        DMinLib = rs!results
    End If
    
    rs.Close
    db.Close
    
    Exit Function
Function_Err:
    Select Case err.Number
        Case 3265:
            Resume Next
        Case Else:
    End Select
End Function

Public Function dLookUpLib(Campo, Tabla, Optional Criterio, Optional ORDEN)
    Dim bd As Database
    Dim rs As Recordset

    Set bd = CurrentDbLib(dll_db_lib)
    
    If Not IsMissing(Criterio) Then
        Criterio = Nz(Criterio)
    Else
        Criterio = ""
    End If
    
    If Not IsMissing(ORDEN) Then
        ORDEN = Nz(ORDEN)
    Else
        ORDEN = ""
    End If
   
    If Left(Tabla, 1) <> "[" Then
        Tabla = "[" & Tabla & "]"
    End If
    
    Set rs = bd.OpenRecordset("SELECT " & Campo & " as Result FROM " & Tabla & IIf(Criterio <> "", " WHERE " & Criterio, "") & IIf(ORDEN <> "", " ORDER BY " & ORDEN, ""), dbOpenSnapshot)
    If Not rs.EOF Then
        dLookUpLib = rs!result
    Else
        dLookUpLib = Null
    End If
    rs.Close
    bd.Close
    
End Function


Public Function dFirstLib(Campo, Tabla, Optional Criterio, Optional ORDEN)
    Dim bd As Database
    Dim rs As Recordset

    Set bd = CurrentDbLib
    
    If Not IsMissing(Criterio) Then
        Criterio = Nz(Criterio)
    Else
        Criterio = ""
    End If
    
    If Not IsMissing(ORDEN) Then
        ORDEN = Nz(ORDEN)
    Else
        ORDEN = ""
    End If
   
    Set rs = bd.OpenRecordset("SELECT first(" & Campo & ") as Ret FROM " & Tabla & IIf(Criterio <> "", " WHERE " & Criterio, "") & IIf(ORDEN <> "", " ORDER BY " & ORDEN, ""), dbOpenSnapshot)
    
    If Not rs.EOF Then
        dFirstLib = rs("Ret")
    Else
        dFirstLib = Null
    End If

End Function

