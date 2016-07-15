' This is a part of the VBto Converter (www.vbto.net). Copyright (C) 2005-2011 StressSoft Company Ltd. All rights reserved.

Imports System.Data.OleDb
Imports System.ComponentModel

Module VBto_Converter

    Public Sub LoadUnUsed(ByVal frm As Control)
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
    End Sub
    Public Sub LoadUnUsed(ByVal frm As MenuItem)
    End Sub


    ' === External Consts: ===
    Public Const Color As Integer = 4

End Module	' VBto_Converter

Namespace VBto

    '============================================================
    ' This is a part of the VBto Converter 2.46
    ' Copyright (C) 2005-2011 StressSoft Company Ltd. All rights reserved
    ' http://www.vbto.net
    '============================================================
    Public Class VBtoRecordSet
        Private dataAdapter As New OleDbDataAdapter()

        Public ReadOnly Property Index() As Integer
            Get
                Return Me.Binding.Position
            End Get
        End Property

        Public ReadOnly Property RowNo() As Integer
            Get
                Return Me.Index + 1
            End Get
        End Property

        Private dataTable As New DataTable()
        Public ReadOnly Property Table() As DataTable
            Get
                Return Me.dataTable
            End Get
        End Property

        Private FActiveConnection As New OleDbConnection()
        Public Property ActiveConnection() As OleDbConnection
            Get
                Return Me.FActiveConnection
            End Get
            Set(ByVal value As OleDbConnection)
                Me.FActiveConnection = value
            End Set
        End Property

        Public Property CommandType() As CommandType
            Get
                Return Me.FActiveCommand.CommandType
            End Get
            Set(ByVal value As CommandType)
                Me.FActiveCommand.CommandType = value
            End Set
        End Property

        Private FActiveCommand As New OleDbCommand()
        Public Property ActiveCommand() As OleDbCommand
            Get
                Return Me.FActiveCommand
            End Get
            Set(ByVal value As OleDbCommand)
                Me.FActiveCommand = value
            End Set
        End Property

        Private FCurrentIndex As Integer = 0

        Public Property AbsolutePosition() As Integer
            Get
                Return Me.FBinding.Position
            End Get
            Set(ByVal value As Integer)
                Me.FBinding.Position = value
            End Set
        End Property

        Public ReadOnly Property EditMode() As DataRowState
            Get
                Return Me.Fields.Row.RowState
            End Get
        End Property

        Private FBinding As New BindingSource()
        Public ReadOnly Property Binding() As BindingSource
            Get
                Return Me.FBinding
            End Get
        End Property

        Public Sub AddNew()
            Me.Binding.AddNew()
        End Sub

        Private FBOF As [Boolean] = True
        Public ReadOnly Property BOF() As [Boolean]
            Get
                Return Me.FBOF
            End Get
        End Property

        Private FEOF As [Boolean] = False
        Public ReadOnly Property EOF() As [Boolean]
            Get
                Return Me.FEOF
            End Get
        End Property

        Public ReadOnly Property Fields() As DataRowView
            Get
                Return DirectCast(Me.Binding.Current, DataRowView)
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As Integer) As [Object]
            Get
                Return Me.Fields(index)
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal [property] As String) As [Object]
            Get
                Return Me.Fields([property])
            End Get
        End Property

        Public ReadOnly Property ColumnCount() As Integer
            Get
                Return Me.dataTable.Columns.Count
            End Get
        End Property

        Public ReadOnly Property Column() As DataColumnCollection
            Get
                Return Me.dataTable.Columns
            End Get
        End Property

        Private Function ValidateSQLText(ByVal value As String) As Boolean
            Dim SQL As Boolean = False
            Dim sqlwords As String() = New String() {"select", "insert", "update", "delete"}
            For Each word As String In sqlwords
                If value.ToLower().IndexOf(word) <> -1 Then
                    SQL = True
                    Exit For
                End If
            Next
            Return SQL
        End Function

        Public Property RecordSource() As String
            Get
                Return FActiveCommand.CommandText
            End Get
            Set(ByVal value As String)
                If Me.ValidateSQLText(value) Then
                    Me.FActiveCommand.CommandText = value
                Else
                    Me.FActiveCommand.CommandText = String.Format("select * from {0}", value)
                End If
                Me.dataTable.Clear()
                Me.dataTable.Columns.Clear()
            End Set
        End Property

#Region "Open methods..."

        Public Sub Open()
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If
            If Me.ActiveConnection.ConnectionString.Length <> 0 Then
                Dim oledbconnBuilder As New OleDbConnectionStringBuilder(Me.ActiveConnection.ConnectionString)
                If System.IO.File.Exists(oledbconnBuilder.DataSource) OrElse System.IO.Directory.Exists(oledbconnBuilder.DataSource) Then
                    Me.FActiveConnection.Open()
                    Me.FActiveCommand.Connection = Me.FActiveConnection
                    Me.dataAdapter.SelectCommand = Me.FActiveCommand
                    Me.dataTable.Clear()
                    Me.dataAdapter.Fill(Me.dataTable)

                    Me.FBinding.DataSource = Me.dataTable
                End If
            End If
        End Sub

        Public Sub Open(ByVal Source As OleDbCommand)
            Me.FActiveCommand = Source
            Me.Open()
        End Sub

        Public Sub Open(ByVal Source As [String])
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If

            Me.RecordSource = Source
            Me.Open()
        End Sub

        Public Sub Open(ByVal Source As [String], ByVal ActiveConnect As OleDbConnection)
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If

            Me.FActiveConnection = ActiveConnect
            Me.Open(Source)
        End Sub

        Public Sub Open(ByVal Source As [String], ByVal ActiveConnect As [String])
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If

            Me.FActiveConnection.ConnectionString = ActiveConnect
            Me.Open(Source)
        End Sub

        Public Sub Open(ByVal Source As OleDbCommand, ByVal ActiveConnect As OleDbConnection)
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If

            Me.FActiveConnection = ActiveConnect
            Me.FActiveCommand = Source
            Me.Open()
        End Sub

        Public Sub Open(ByVal Source As OleDbCommand, ByVal ActiveConnect As [String])
            If Me.FActiveConnection.State = ConnectionState.Open Then
                Me.FActiveConnection.Close()
            End If

            Me.FActiveConnection.ConnectionString = ActiveConnect
            Me.FActiveCommand = Source
            Me.Open()
        End Sub

#End Region

        Public Sub Close()
            Me.ActiveConnection.Close()
        End Sub

        Public Sub Refresh()
            Me.Open()
        End Sub

        Public Sub Edit()
            Me.Fields.BeginEdit()
        End Sub

        Public Enum MoveStat
            adBookmarkCurrent = 0
            adBookmarkFirst = 1
            adBookmarkLast = 2
        End Enum

        Private Sub UpdateBofAndEof()
            Me.FBOF = (Me.Binding.Position = 0)
            Me.FEOF = (Me.Binding.Position = Me.Binding.Count - 1)
        End Sub

        Public Sub Update()
            Using FCommandBuilder As New OleDbCommandBuilder(dataAdapter)
                Me.dataAdapter.UpdateCommand = FCommandBuilder.GetUpdateCommand()
                Me.dataAdapter.InsertCommand = FCommandBuilder.GetInsertCommand()
                Me.dataAdapter.DeleteCommand = FCommandBuilder.GetDeleteCommand()
                Try
                    Me.Binding.EndEdit()
                    Me.dataAdapter.Update(Me.dataTable)
                Catch e As Exception
                    MessageBox.Show(e.Message)
                End Try
            End Using
        End Sub

        Public Sub Delete()
            Me.Fields.Delete()
            Me.Update()
        End Sub

#Region "Move methods..."

        Public Sub Move(ByVal NumRecords As Int32, ByVal Stat As MoveStat)
            Me.FCurrentIndex = Me.Binding.Position
            Select Case Stat
                Case MoveStat.adBookmarkFirst
                    If NumRecords >= dataTable.Rows.Count Then
                        Me.FCurrentIndex = Me.dataTable.Rows.Count - 1
                    Else
                        If NumRecords > 0 Then
                            Me.FCurrentIndex = NumRecords
                        Else
                            Me.FCurrentIndex = Me.dataTable.Rows.Count + NumRecords
                        End If
                    End If
                    Exit Select
                Case MoveStat.adBookmarkLast
                    If NumRecords >= Me.dataTable.Rows.Count Then
                        Me.FCurrentIndex = 0
                    Else
                        If NumRecords > 0 Then
                            Me.FCurrentIndex = Me.dataTable.Rows.Count - NumRecords - 1
                        Else
                            Me.FCurrentIndex = NumRecords
                        End If
                    End If
                    Exit Select
                Case Else
                    If NumRecords = Me.dataTable.Rows.Count Then
                        Me.FCurrentIndex = Me.dataTable.Rows.Count
                    Else
                        Me.FCurrentIndex = NumRecords
                    End If
                    Exit Select
            End Select

            If Me.EOF AndAlso Not Me.BOF Then
                Me.Binding.Position = Me.FCurrentIndex - 1
            ElseIf Not Me.EOF AndAlso Not Me.BOF Then
                If Me.FCurrentIndex > 0 AndAlso Me.FCurrentIndex < Me.dataTable.Rows.Count Then
                    Me.Binding.Position = Me.FCurrentIndex
                End If
            ElseIf Not Me.EOF AndAlso Me.BOF Then
                Me.Binding.Position = Me.FCurrentIndex + 1
            End If

            Me.UpdateBofAndEof()
        End Sub

        Public Sub Move(ByVal NumRecords As Int32)
            Me.UpdateBofAndEof()
            Me.Move(NumRecords, MoveStat.adBookmarkCurrent)
        End Sub

        Public Sub MoveFirst()
            Me.Binding.MoveFirst()
            Me.UpdateBofAndEof()
        End Sub

        Public Sub MoveLast()
            Me.Binding.MoveLast()
            Me.UpdateBofAndEof()
        End Sub

        Public Sub MoveNext()
            Me.UpdateBofAndEof()
            Me.Binding.MoveNext()
        End Sub

        Public Sub MovePrevious()
            Me.UpdateBofAndEof()
            Me.Binding.MovePrevious()
        End Sub

#End Region

        Public ReadOnly Property RecordCount() As Integer
            Get
                Return Me.FBinding.Count
            End Get
        End Property

#Region "Find..."

        Private Class Finder
            Private FRows As DataRow() = New DataRow(-1) {}
            Private FTable As DataTable
            Private FindWord As String
            Private FindField As String
            Private CurrentPos As Integer = -1
            Private FUniqueField As String = ""

            Public ReadOnly Property UniqueField() As String
                Get
                    Return Me.FUniqueField
                End Get
            End Property

            Private Function GetPrimaryKeys(ByVal conn As OleDbConnection) As String
                Dim schemaTable As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, New Object() {Nothing, Nothing, Me.FTable.TableName})
                Return schemaTable.Rows(0).ItemArray(3).ToString()
            End Function

            Public Function FindFirst(ByVal table As DataTable, ByVal FieldName As String, ByVal FindValue As String, ByVal conn As OleDbConnection) As DataRow
                Dim flag As Boolean = False

                If Me.FTable IsNot Nothing Then
                    If Me.FTable Is table Then
                        flag = True
                        If Me.FindField.Equals(FieldName) Then
                            flag = Me.FindWord.Equals(FindValue)
                        Else
                            flag = False
                        End If
                    End If
                End If

                If flag Then
                    If Me.FRows.Length > 0 Then
                        For Each Row As DataRow In Me.FRows
                            If Row.RowState = DataRowState.Detached Then
                                flag = False
                                Exit For
                            End If
                        Next
                    End If
                End If

                If Not flag Then
                    Me.FTable = table
                    Me.FindField = FieldName
                    Me.FUniqueField = Me.FindField
                    Me.FindWord = FindValue
                    Me.CurrentPos = -1

                    Array.Clear(Me.FRows, 0, Me.FRows.Length)

                    If Me.FTable.Columns(Me.FindField).DataType Is System.Type.[GetType]("System.String") Then
                        Me.FRows = Me.FTable.[Select](Me.FindField & " Like '" & Me.FindWord & "'")
                    Else
                        Me.FRows = Me.FTable.[Select](Me.FindField & " = " & Me.FindWord)
                    End If

                    If Me.FRows.Length > 0 Then
                        Me.CurrentPos = 0

                        ' set UniqueField
                        If Me.FTable.TableName.Length > 0 Then
                            Me.FUniqueField = GetPrimaryKeys(conn)
                            'foreach (DataColumn cl in FTable.Columns)
                            '{
                            '    var query = from rows in FRows
                            '                group rows by rows.Field<object>(cl.Caption) into g
                            '                select new { ValueCount = g.Count() };

                            '    if (query.Count(x => x.ValueCount > 1) == 0)
                            '    {
                            '        this.FUniqueField = cl.Caption;
                            '        break;
                            '    }
                            '}
                        End If
                    End If
                End If

                If Me.CurrentPos = -1 Then
                    Return Nothing
                End If

                Dim rw As DataRow = Me.FRows(System.Math.Max(System.Threading.Interlocked.Increment(Me.CurrentPos), Me.CurrentPos - 1))

                If Me.CurrentPos >= Me.FRows.Length Then Me.CurrentPos = 0

                Return rw
            End Function
        End Class

        Private FindObj As New Finder()

        Public Function FindFirst(ByVal [property] As String, ByVal value As String) As Integer
            Dim row As DataRow = FindObj.FindFirst(Me.Table, [property], value, Me.ActiveConnection)
            If row Is Nothing Then
                Return -1
            End If
            Dim pdc As PropertyDescriptorCollection = Me.FBinding.CurrencyManager.GetItemProperties()
            Dim tmp As Integer = Me.FBinding.Find(pdc(FindObj.UniqueField), row(FindObj.UniqueField))
            If tmp <> -1 Then
                Me.FBinding.Position = tmp
            End If
            Return tmp
        End Function

#End Region
    End Class   ' VBtoRecordSet

End Namespace	' VBto
