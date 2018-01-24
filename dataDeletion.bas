Attribute VB_Name = "dataDeletion"
Option Compare Database
Option Explicit

'====================================================
'---------------DATA Deletion FUNCTIONS----------------
'Added a new comment
'Added some new comment since now i am working on a branch
'Added further comments
'====================================================

Public Function Data_exists(firmID As Integer, year As Integer, tblName As String) As Boolean

Dim strSQL As String
Dim rsData As DAO.Recordset

    strSQL = "select * from " & tblName & " where firm_id = " & firmID _
            & "and report_year = " & year
    
    Set rsData = CurrentDb.OpenRecordset(strSQL)
    If rsData.RecordCount = 0 Then
        Data_exists = False
    Else
        Data_exists = True
    End If
    
    rsData.Close

End Function

Public Sub DeleteData(firmID As Integer, year As Integer, tblName As String)

Dim strSQL As String
Dim dbGDC As Database


    Set dbGDC = CurrentDb
    
    strSQL = "DELETE * FROM " & tblName _
           & " WHERE firm_id  = " & firmID _
           & " AND report_year  = " & year

    dbGDC.Execute strSQL, dbFailOnError
    Debug.Print strSQL

End Sub


Public Sub DeleteFirmData(firmID As Integer)

Dim strSQL As String
Dim dbGDC As Database


    Set dbGDC = CurrentDb
    
    strSQL = "DELETE * FROM Firms WHERE ID  = " & firmID
           
    'dbGDC.Execute strSQL, dbFailOnError
    Debug.Print strSQL

End Sub
