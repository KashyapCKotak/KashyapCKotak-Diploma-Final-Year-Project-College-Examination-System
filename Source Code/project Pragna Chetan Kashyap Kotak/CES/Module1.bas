Attribute VB_Name = "Module1"
Global con As ADODB.Connection
Global rs As ADODB.Recordset
Global uname As String
Global bcode As String
Global sem As Integer
Global subcode As String
Global exid As String
Global t As Integer
Global sh As Integer
Global data As String
Global time As Integer
Global abc As Integer
Global cnt As Integer
Global color As Long
Global xyz As Integer
Global str As String
Global examdecide As Integer
Global timestr2 As String
Global i As Double
Global time2 As Integer
Global concat As Integer
Global counter As Integer
Global time3 As String

Public Function connectdb()
Set con = New ADODB.Connection
'con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\OfflineExaminer.mdb;Persist Security Info=False")

'str = "\\192.168.1.194\kashyap\offlineExaminer.mdb"
'con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & str & ";Persist Security Info=False")
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & str & "\Examiner.mdb;Persist Security Info=False")
'str = "\\ADMIN-PC\New folder\OfflineExaminer.mdb"


End Function
