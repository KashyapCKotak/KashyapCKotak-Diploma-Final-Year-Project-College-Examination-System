'Option Strict Off
'Option Explicit On
Imports VB = Microsoft.VisualBasic

Imports System.Data.OleDb
Imports VBto

Public Module Module1


	'=========================================================
    Public con As OleDbConnection
    Public rs As VBtoRecordSet
    Public uname As String
    Public bcode As String
    ' VBto upgrade warning: sem As Short	OnWrite(String)
    Public sem As Short
    Public subcode As String
    Public exid As String
    ' VBto upgrade warning: t As Short	OnWrite(VBtoRecordSet)
    Public t As Short
    Public sh As Short
    Public data As String
    Public time As Short
    Public abc As Short
    ' VBto upgrade warning: cnt As Short	OnWrite(VBtoRecordSet, Short)
    Public cnt As Short
    Public color As Integer
    Public xyz As Short
    Public str As String
    Public examdecide As Short
    Public timestr2 As String
    Public i As Double
    Public time2 As Short
    Public concat As Short
    Public counter As Short
    Public time3 As String

    Public Function connectdb() As Object
        connectdb = 0
'#Const def_connectdb = True
#If def_connectdb
        con = New OleDbConnection
        'con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\OfflineExaminer.mdb;Persist Security Info=False")

        'str = "\\192.168.1.194\kashyap\offlineExaminer.mdb"
        'con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & str & ";Persist Security Info=False")
        con.Open(("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & str & "\Examiner.mdb;Persist Security Info=False"))
        'str = "\\ADMIN-PC\New folder\OfflineExaminer.mdb"


#End If	' def_connectdb
    End Function

End Module