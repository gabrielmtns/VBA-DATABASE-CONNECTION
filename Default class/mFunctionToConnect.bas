Attribute VB_Name = "mFunctionTConnect"

Option Explicit
Option Private Module

Public Function app(Optional ByVal sql As String, Optional ByVal getRecordset As Boolean) As clsConnection
    Dim myObject As New clsConnection
    Dim dbPath   As String

    dbPath = 'SUA STRING DE CONEXÃO AQUI'

    myObject.connectionString = dbPath
                           
    If Not sql = vbNullString Then myObject.sql = sql
    myObject.create getRecordset
    
    Set app = myObject
End Function

