'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Data.Exceptions
Imports Documents.SerializationUtilities
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Data
  ''' <summary>Defines a data source for use in database lookups and searches.</summary>
  <XmlInclude(GetType(DataMap))>
  Public Class DataSource
    Inherits DataLookup
    Implements IDisposable
    Implements ISerialize
    Implements IPropertyLookup
    'Implements ILoggable

#Region "Class Enumerations"

    Public Enum DataSourceType
      Table = 1
      View = 2
      ProviderTypes = GET_PROVIDER_TYPES
    End Enum

#End Region

#Region "Class Constants"

    Public Const GET_PROVIDER_TYPES As Integer = -100
    ''' <summary>
    ''' Constant value representing the 
    ''' file extension to save XML 
    ''' serialized instances to.
    ''' </summary>
    Public Const DATA_SOURCE_FILE_EXTENSION As String = "dsf"

#End Region

#Region "Class Variables"

    Private mstrConnectionString As String
    Private mstrUserName As String
    Private Password As String
    'Private mobjCriteria As New Criteria
    Private mstrSource As String
    Private mstrDelimiter As String = Nothing
    Private mstrQueryTarget As String
    'Private mobjDocument As Document
    Private mobjConnection As IDbConnection
    Private mobjQueryTargets As New QueryTargets
    Private mstrSQLStatement As String = ""
    Private WithEvents MobjResultColumns As New ResultColumns
    Private mIntLimitResults As Integer
    Private mobjOrderBy As New OrderItems
    Private mobjSearchType As SearchType
    Private mstrContentKeywords As String
    Private mboolDistinct As Boolean
    Private mobjClauses As New Clauses
    Private mobjSearch As Providers.CSearch
    Private mobjRemovedColumns As New List(Of String)

#End Region

#Region "Public Properties"

    Friend ReadOnly Property RemovedColumns As List(Of String)
      Get
        Return mobjRemovedColumns
      End Get
    End Property

    Public Property ContentKeywords() As String
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mstrContentKeywords

      End Get
      Set(ByVal Value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mstrContentKeywords = Value

      End Set
    End Property

    Public Property Search() As Providers.CSearch
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mobjSearch

      End Get
      Set(ByVal Value As Providers.CSearch)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mobjSearch = Value

      End Set
    End Property

    Public Property SearchType() As SearchType
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mobjSearchType

      End Get
      Set(ByVal Value As SearchType)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mobjSearchType = Value

      End Set
    End Property

    Public Property OrderBy() As OrderItems
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mobjOrderBy

      End Get
      Set(ByVal Value As OrderItems)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mobjOrderBy = Value

      End Set
    End Property

    Public Property LimitResults() As Integer
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mIntLimitResults

      End Get
      Set(ByVal Value As Integer)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mIntLimitResults = Value

      End Set
    End Property

    Public Property DistinctQuery() As Boolean
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mboolDistinct

      End Get
      Set(ByVal Value As Boolean)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mboolDistinct = Value

      End Set
    End Property

    Public Property ConnectionString() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          'ApplicationLogging.WriteLogEntry("Enter DataSource::Get_ConnectionString", TraceEventType.Verbose)
          'ApplicationLogging.WriteLogEntry("Exit DataSource::Get_ConnectionString", TraceEventType.Verbose)
          Return mstrConnectionString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, "ExplorerHandlers::Get_ConnectionString")
          Return ""
        End Try
      End Get
      Set(ByVal Value As String)
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          'ApplicationLogging.WriteLogEntry("Enter DataSource::Set_ConnectionString", TraceEventType.Verbose)
          mstrConnectionString = Value

          Try
            mobjQueryTargets = GetTablesAndViews()
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          End Try

        Catch ex As Exception
          ApplicationLogging.LogException(ex, "ExplorerHandlers::Set_ConnectionString")
          Throw New InvalidConnectionStringException(ex.Message, ex)
        Finally
          'ApplicationLogging.WriteLogEntry("Exit DataSource::Set_ConnectionString", TraceEventType.Verbose)
        End Try
      End Set
    End Property

    Public Property QueryTarget() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          Return mstrQueryTarget
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As String)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          mstrQueryTarget = Value.Trim
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property SourceColumn() As String
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mstrSource

      End Get
      Set(ByVal Value As String)

        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          If mstrSource IsNot Nothing AndAlso mstrSource.Length > 0 AndAlso mstrSource <> Value Then
            If ResultColumns.Contains(mstrSource) = True Then
              ResultColumns.Remove(mstrSource)
            End If
          End If

          mstrSource = Value

          If ResultColumns.Contains(Value) = False Then
            If mobjRemovedColumns.Contains(Value) = False Then
              ResultColumns.Insert(0, Value)
            End If
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Set

    End Property

    Public Property Delimiter() As String
      Get
        Return mstrDelimiter
      End Get
      Set(ByVal value As String)
        mstrDelimiter = value
      End Set
    End Property

    Public Property ResultColumns() As ResultColumns
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return MobjResultColumns

      End Get
      Set(ByVal value As ResultColumns)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        MobjResultColumns = value

      End Set
    End Property

    Friend Property Criteria() As Criteria
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          Return Clauses.GetFirstClause.Criteria 'mobjCriteria
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As Criteria)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          Clauses.GetFirstClause.Criteria = Value 'mobjCriteria = Value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Clauses() As Clauses
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mobjClauses

      End Get
      Set(ByVal value As Clauses)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mobjClauses = value

      End Set
    End Property

    Public Property SQLStatement() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          If Helper.CallStackContainsMethodName("Serialize") = False Then
            If mstrSQLStatement Is Nothing OrElse mstrSQLStatement.Length = 0 Then
              Throw New InvalidOperationException("The SQLStatement has not yet been initialized.  Call SQLStatement(lpDocument) or SQLStatement(lpVersion) instead.")
            End If
          End If

          Return mstrSQLStatement
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mstrSQLStatement = value

      End Set
    End Property

#End Region

#Region "Public ReadOnly Properties"

    'Public Property Document() As Document
    '  Get
    '    Return mobjDocument
    '  End Get
    '  Set(ByVal Value As Document)
    '    mobjDocument = Value
    '  End Set
    'End Property

    Public ReadOnly Property ResultColumnsString() As String
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return BuildResultColumnsString()

      End Get
    End Property

    Public ReadOnly Property Provider() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          Return GetInfoFromString(ConnectionString, "Provider")

        Catch ex As Exception
          ApplicationLogging.LogException(ex, "DataSource::Get_Provider")
          Return ""
        End Try
      End Get
    End Property

    Public ReadOnly Property SQLStatement(ByVal lpMetaHolder As Core.IMetaHolder) As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          mstrSQLStatement = BuildSQLString(lpMetaHolder)

          Return mstrSQLStatement

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return ""
        End Try
      End Get
    End Property

    'Public ReadOnly Property SQLStatement(ByVal lpDocument As Core.Document) As String
    '  Get
    '    Try
    '      mstrSQLStatement = BuildSQLString(lpDocument)
    '      Return mstrSQLStatement
    '    Catch ex As Exception
    '       ApplicationLogging.LogException(ex, "DataSource::Get_SQLStatement(lpDocument)")
    '      Return ""
    '    End Try
    '  End Get
    'End Property

    'Public ReadOnly Property SQLStatement(ByVal lpVersion As Core.Version) As String
    '  Get
    '    Try
    '      mstrSQLStatement = BuildSQLString(lpVersion)
    '      Return mstrSQLStatement
    '    Catch ex As Exception
    '       ApplicationLogging.LogException(ex, "DataSource::Get_SQLStatement(lpVersion)")
    '      Return ""
    '    End Try
    '  End Get
    'End Property

    <XmlIgnore()>
    Public ReadOnly Property QueryTargets() As QueryTargets
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return mobjQueryTargets
        'Return GetTablesAndViews()

      End Get
    End Property

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Friend Properties"

    Friend WriteOnly Property SearchClauses() As Search.Clauses
      Set(ByVal value As Search.Clauses)

        ' Clear the existing clauses
        Clauses.Clear()

        ' Translate to the data clauses
        For Each lobjClause As Search.Clause In value
          Clauses.Add(lobjClause)
        Next

      End Set
    End Property

    <XmlIgnore()>
    Public Property Connection() As IDbConnection
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


          If mobjConnection IsNot Nothing Then
            Return mobjConnection
          Else
            InitializeConnection()
            Return mobjConnection
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, "DataSource::Get_Connection")
          Return Nothing
        End Try
      End Get
      Set(ByVal Value As IDbConnection)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        mobjConnection = Value

      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(LookupType.Map)
      Clauses.Add(New Clause())
    End Sub

    Public Sub New(ByVal lpSearch As Providers.CSearch)
      MyBase.New(LookupType.Map)
      Clauses.Add(New Clause())
    End Sub

    Public Sub New(ByVal lpConnectionString As String,
                   ByVal lpQueryTarget As String,
                   ByVal lpSourceColumn As String,
                   ByVal lpCriteria As Criteria)

      MyBase.New(LookupType.Map)

      Try
        ConnectionString = lpConnectionString
        QueryTarget = lpQueryTarget
        SourceColumn = lpSourceColumn
        ResultColumns.Add(lpSourceColumn)
        Criteria = lpCriteria
        Clauses.Add(New Clause())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::New(lpConnectionString:{0}, lpQueryTarget:{1}, lpSourceColumn:{2}, lpCriteria", lpConnectionString, lpQueryTarget, lpSourceColumn))
      End Try

    End Sub

    Public Sub New(ByVal lpXMLFilePath As String)

      MyBase.New(LookupType.Map)

      Try
        Dim lobjDataSource As DataSource = Deserialize(lpXMLFilePath)
        With Me
          .ConnectionString = lobjDataSource.ConnectionString
          .QueryTarget = lobjDataSource.QueryTarget
          .SourceColumn = lobjDataSource.SourceColumn
          .ResultColumns.Add(lobjDataSource.SourceColumn)
          .Criteria = lobjDataSource.Criteria
          .mstrUserName = lobjDataSource.mstrUserName
          .Password = lobjDataSource.Password
          .Clauses = lobjDataSource.Clauses
        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::New(lpXMLFilePath:{0}", lpXMLFilePath))
      End Try

    End Sub

#End Region

#Region "Public Methods"

    Public Sub Reset()

      Try
        Me.QueryTarget = String.Empty
        Me.ResultColumns.Clear()
        Me.OrderBy.Clear()
        Me.Clauses.Clear()
        Me.Clauses.Add(New Data.Clause())
        Me.RemovedColumns.Clear()
        Me.DistinctQuery = False
        Me.LimitResults = 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Function GetColumns(ByVal lpFilter As String) As ArrayList
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Dim lobjArrayList As New ArrayList
        Dim lobjDataTable As System.Data.DataTable
        Dim lstrObjectName As String
        'LogSession.LogVerbose("GetColumns({0}): About to call GetSchema({1}, {0})", lpFilter, GET_PROVIDER_TYPES)
        lobjDataTable = GetSchema(GET_PROVIDER_TYPES, lpFilter)
        'LogSession.LogVerbose("GetColumns({0}): Completed call to GetSchema({1}, {0})", lpFilter, GET_PROVIDER_TYPES)
        For Each lobjDataRow As System.Data.DataRow In lobjDataTable.Rows
          lstrObjectName = lobjDataRow.Item("COLUMN_NAME")

          If Not lstrObjectName.ToUpper.StartsWith("SYS") Then
            If Not lstrObjectName.ToUpper.StartsWith("MSYS") Then
              If Not lobjArrayList.Contains(lstrObjectName) Then
                lobjArrayList.Add(lstrObjectName)
                Debug.WriteLine(lstrObjectName)
              End If
            End If
          End If
          'lobjQueryTargets.Add(New QueryTarget(lobjDataRow.Item(2), QueryTargetType.Table))

        Next

        Return lobjArrayList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    'Public Function GetColumns(ByVal lpFilter As String, ByVal lpDataType As Nullable(Of OleDb.OleDbType)) As ArrayList
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Dim lobjArrayList As New ArrayList
    '    Dim lobjDataTable As System.Data.DataTable
    '    Dim lstrObjectName As String

    '    'LogSession.LogVerbose("GetColumns({0}, {1}): About to call GetSchema({2}, {0})", lpFilter, lpDataType, GET_PROVIDER_TYPES)
    '    lobjDataTable = GetSchema(GET_PROVIDER_TYPES, lpFilter)
    '    'LogSession.LogVerbose("GetColumns({0}, {1}): Completed call to GetSchema({2}, {0})", lpFilter, lpDataType, GET_PROVIDER_TYPES)
    '    If lpDataType.HasValue Then
    '      Dim lenuDataType As OleDb.OleDbType
    '      For Each lobjDataRow As System.Data.DataRow In lobjDataTable.Rows
    '        lstrObjectName = lobjDataRow.Item("COLUMN_NAME")
    '        lenuDataType = lobjDataRow.Item("DATA_TYPE")
    '        If Not lstrObjectName.ToUpper.StartsWith("SYS") Then
    '          If Not lstrObjectName.ToUpper.StartsWith("MSYS") Then
    '            If Not lobjArrayList.Contains(lstrObjectName) Then
    '              If lpDataType = lenuDataType Then
    '                lobjArrayList.Add(lstrObjectName)
    '                Debug.WriteLine(lstrObjectName)
    '              End If
    '            End If
    '          End If
    '        End If
    '        'lobjQueryTargets.Add(New QueryTarget(lobjDataRow.Item(2), QueryTargetType.Table))

    '      Next
    '    Else
    '      For Each lobjDataRow As System.Data.DataRow In lobjDataTable.Rows
    '        lstrObjectName = lobjDataRow.Item("COLUMN_NAME")

    '        If Not lstrObjectName.ToUpper.StartsWith("SYS") Then
    '          If Not lstrObjectName.ToUpper.StartsWith("MSYS") Then
    '            If Not lobjArrayList.Contains(lstrObjectName) Then
    '              lobjArrayList.Add(lstrObjectName)
    '              Debug.WriteLine(lstrObjectName)
    '            End If
    '          End If
    '        End If
    '        'lobjQueryTargets.Add(New QueryTarget(lobjDataRow.Item(2), QueryTargetType.Table))

    '      Next
    '    End If


    '    Return lobjArrayList

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Function

    Public Function GetTablesAndViews() As QueryTargets

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        'ApplicationLogging.WriteLogEntry("Enter DataSource::GetTablesAndViews", TraceEventType.Verbose)
        Dim lobjQueryTargets As New QueryTargets
        Dim lobjDataTable As System.Data.DataTable
        Dim lblnHasTableSchema As Boolean
        Dim lstrTableSchema As String = String.Empty
        Dim lstrObjectName As String = String.Empty
        Dim lstrObjectType As String = String.Empty
        Dim lstrQualifiedTableName As String = String.Empty

        ' Get the 'Tables'
        'LogSession.LogVerbose("GetTablesAndViews: About to call GetSchema({0})", DataSourceType.Table.ToString())
        lobjDataTable = GetSchema(DataSourceType.Table)
        'LogSession.LogVerbose("GetTablesAndViews: Completed call to GetSchema({0})", DataSourceType.Table.ToString())
        lblnHasTableSchema = lobjDataTable.Columns.Contains("TABLE_SCHEMA")

        For Each lobjDataRow As System.Data.DataRow In lobjDataTable.Rows
          If lblnHasTableSchema AndAlso Not IsDBNull(lobjDataRow.Item("TABLE_SCHEMA")) Then
            lstrTableSchema = lobjDataRow.Item("TABLE_SCHEMA")
          Else
            lstrTableSchema = String.Empty
          End If
          lstrObjectName = lobjDataRow.Item("TABLE_NAME")
          lstrObjectType = lobjDataRow.Item("TABLE_TYPE")
          Debug.WriteLine(String.Format("{0}: {1}", lstrObjectName, lstrObjectType))
          If Not lobjQueryTargets.Contains(lstrObjectName) Then
            If lblnHasTableSchema Then
              If Not String.IsNullOrEmpty(lstrTableSchema) Then
                lstrQualifiedTableName = String.Format("{0}.{1}", lstrTableSchema, lstrObjectName)
              Else
                lstrQualifiedTableName = lstrObjectName
              End If
              lobjQueryTargets.Add(New QueryTarget(lstrObjectName, lstrQualifiedTableName, QueryTargetType.Table))
            Else
              lobjQueryTargets.Add(New QueryTarget(lstrObjectName, QueryTargetType.Table))
            End If
          End If
        Next

        ' Get the 'Views'
        Try
          'LogSession.LogVerbose("GetTablesAndViews: About to call GetSchema({0})", DataSourceType.View.ToString())
          lobjDataTable = GetSchema(DataSourceType.View)
          'LogSession.LogVerbose("GetTablesAndViews: Completed call to GetSchema({0})", DataSourceType.View.ToString())
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          If ex.Message.StartsWith("The Views OleDbSchemaGuid is not a supported schema") Then
            ' We can't get the 'Views'
            Return lobjQueryTargets
          End If
        End Try

        For Each lobjDataRow As System.Data.DataRow In lobjDataTable.Rows
          lstrObjectName = lobjDataRow.Item("TABLE_NAME")
          If Not lobjQueryTargets.Contains(lstrObjectName) Then
            lobjQueryTargets.Add(New QueryTarget(lstrObjectName, QueryTargetType.View))
          End If
        Next

        Return lobjQueryTargets
        'lobjOleDbConnection.Close()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "DataSource::GetTablesAndViews")
        ' Re-throw the exception to the caller
        Throw
      Finally
        'ApplicationLogging.WriteLogEntry("Exit DataSource::GetTablesAndViews", TraceEventType.Verbose)
      End Try

    End Function

    Public Function GetSchema(ByVal lpDataSourceType As DataSourceType, lpOwner As String, Optional ByVal lpFilter As String = "") As System.Data.DataTable

      Dim lobjSchemaGuid As Guid
      Dim lobjRestrictions As Object()

      Try

        If String.IsNullOrEmpty(ConnectionString) Then
          Throw New InvalidOperationException()
        End If

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        If String.IsNullOrEmpty(lpOwner) Then
          lpOwner = Nothing
        End If

        Select Case lpDataSourceType
          Case DataSourceType.Table
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Tables
            lobjRestrictions = New Object() {Nothing, lpOwner, Nothing, "TABLE"}

          Case DataSourceType.View
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Tables
            lobjRestrictions = New Object() {Nothing, lpOwner, Nothing, "VIEW"}

          Case DataSourceType.ProviderTypes
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Columns
            lobjRestrictions = New Object() {Nothing, Nothing, lpFilter}

          Case Else
            Throw New InvalidOperationException("Value '" & lpDataSourceType & "' not supported for operation.")
        End Select

        Dim lobjOleDbConnection As OleDb.OleDbConnection

        Try
          lobjOleDbConnection = New OleDb.OleDbConnection(ConnectionString)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New Exception("Could not create OleDb Connection from connection string '" &
                              ConnectionString & "' [" & ex.Message & "]", ex)
        End Try

        Dim lobjDataTable As System.Data.DataTable

        If lobjOleDbConnection.State = ConnectionState.Closed Then
          lobjOleDbConnection.Open()
        End If

        ' Get the 'Schema'
        'lobjDataTable = lobjOleDbConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
        'LogSession.LogVerbose("GetSchema: About to call OleDbConnection.GetOleDbSchemaTable: ({0}, {1}, {2})", lpDataSourceType.ToString, lpOwner, lpFilter)
        lobjDataTable = lobjOleDbConnection.GetOleDbSchemaTable(lobjSchemaGuid, lobjRestrictions)
        'LogSession.LogVerbose("GetSchema: Completed internal call to OleDbConnection.GetOleDbSchemaTable: ({0}, {1}, {2})", lpDataSourceType.ToString, lpOwner, lpFilter)
        lobjOleDbConnection.Close()

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function GetSchema(ByVal lpDataSourceType As DataSourceType, Optional ByVal lpFilter As String = "") As System.Data.DataTable

      Dim lobjSchemaGuid As Guid
      Dim lobjRestrictions As Object()

      Try

        If String.IsNullOrEmpty(ConnectionString) Then
          Throw New InvalidOperationException()
        End If

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Select Case lpDataSourceType
          Case DataSourceType.Table
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Tables
            lobjRestrictions = New Object() {Nothing, Nothing, Nothing, "TABLE"}

          Case DataSourceType.View
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Tables
            lobjRestrictions = New Object() {Nothing, Nothing, Nothing, "VIEW"}

          Case DataSourceType.ProviderTypes
            lobjSchemaGuid = OleDb.OleDbSchemaGuid.Columns
            lobjRestrictions = New Object() {Nothing, Nothing, lpFilter}

          Case Else
            Throw New InvalidOperationException("Value '" & lpDataSourceType & "' not supported for operation.")
        End Select

        Dim lobjOleDbConnection As OleDb.OleDbConnection

        Try
          lobjOleDbConnection = New OleDb.OleDbConnection(ConnectionString)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New Exception("Could not create OleDb Connection from connection string '" &
                              ConnectionString & "' [" & ex.Message & "]", ex)
        End Try

        Dim lobjDataTable As System.Data.DataTable

        If lobjOleDbConnection.State = ConnectionState.Closed Then
          lobjOleDbConnection.Open()
        End If

        ' Get the 'Schema'
        'lobjDataTable = lobjOleDbConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
        'LogSession.LogVerbose("GetSchema: About to call OleDbConnection.GetOleDbSchemaTable: ({0}, {1})", lpDataSourceType.ToString, lpFilter)
        lobjDataTable = lobjOleDbConnection.GetOleDbSchemaTable(lobjSchemaGuid, lobjRestrictions)
        'LogSession.LogVerbose("GetSchema: Completed internal call to OleDbConnection.GetOleDbSchemaTable: ({0}, {1})", lpDataSourceType.ToString, lpFilter)

        lobjOleDbConnection.Close()

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Executes a database search with the current criteria set using values in the document parameter
    ''' </summary>
    ''' <param name="lpDocument"></param>
    ''' <param name="lpErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataTable(ByVal lpDocument As Core.Document, Optional ByRef lpErrorMessage As String = "") As System.Data.DataTable

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return GetDataTable(SQLStatement(lpDocument), lpErrorMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try

    End Function

    Public Overrides Function GetParameters() As IParameters Implements IPropertyLookup.GetParameters
      Try

        Dim lobjParameters As IParameters = New Core.Parameters

        ' TODO: Get the actual values
        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetDataTable(ByVal lpSQLStatement As String, Optional ByRef lpErrorMessage As String = "") As System.Data.DataTable

      Dim lobjDataTable As New System.Data.DataTable

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        Dim lstrSQL As String = lpSQLStatement.TrimEnd(" ")

#If DEBUG Then
        ' Write the SQL Statement we are about to execute out to the log
        ApplicationLogging.WriteLogEntry(String.Format("About to execute SQL statement '{0}' in {1}", lstrSQL, Reflection.MethodBase.GetCurrentMethod), TraceEventType.Information, 3925)
#End If

        ' Create a query command on the connection
        Dim lobjCommand As New OleDb.OleDbCommand(lstrSQL, Connection)

        ' Run the query; get the DataReader object
        Dim lobjDataReader As OleDb.OleDbDataReader

        Try
          lobjDataReader = lobjCommand.ExecuteReader(CommandBehavior.SingleResult)
        Catch ex As Exception
          lpErrorMessage &= String.Format("An error occured on ExecuteReader in DataSource::GetDataTable with SQL Statement: '{0}'.  The error was '{1}'", lpSQLStatement, ex.Message)
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error)
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          'Throw New Exception("Unable to get results [" & ex.Message & "]", ex)
          Return lobjDataTable
        End Try

        If lobjDataReader.HasRows Then
          ' Read the resultset
          lobjDataTable.Load(lobjDataReader)
          'Do While lobjDataReader.Read
          '  ' Get the first value
          '  'lobjValue = lobjDataReader.Item(Me.SourceColumn)
          '  Exit Do
          'Loop

#If DEBUG Then
          ' Write the SQL Statement we are about to execute out to the log
          ApplicationLogging.WriteLogEntry(String.Format("{0} rows found in {1}", lobjDataTable.Rows.Count, Reflection.MethodBase.GetCurrentMethod), TraceEventType.Information, 3926)
#End If
        Else
          lpErrorMessage = "No value found for the expression (" & lstrSQL & ")"
        End If

        ' Always close the DataReader
        lobjDataReader.Close()

        ' Always close the Connection
        If Connection.State = ConnectionState.Open Then
          Connection.Close()
        End If

        Return lobjDataTable

      Catch ex As Exception
        lpErrorMessage &= String.Format("An error occured on ExecuteReader in DataSource::GetDataTable with SQL Statement: '{0}'.  The error was '{1}'", lpSQLStatement, ex.Message)
        ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error)
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage &= Helper.FormatCallStack(ex)
        Return lobjDataTable
      End Try

    End Function

    ''' <summary>
    ''' Executes a query with the specified SQL statement and returns the first column of the first row.
    ''' </summary>
    ''' <param name="lpSQL"></param>
    ''' <returns>Returns the first column of the first row.  If there are no results then returns an empty string.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteSimpleQuery(ByVal lpSQL As String) As Object

      Dim lobjResult As Object

      Try

        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        ' Create a query command on the connection
        Dim lobjCommand As New OleDb.OleDbCommand(lpSQL, Connection)
        lobjResult = lobjCommand.ExecuteScalar

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally

      End Try

      Return lobjResult

    End Function

#End Region

#Region "Private Methods"

    Protected Overridable Function BuildResultColumnsString() As String
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Dim lstrOutpuString As String = ""
        For Each lstrResultColumn As String In ResultColumns
          If Not String.IsNullOrEmpty(lstrResultColumn) Then
            If lstrResultColumn.Contains(" "c) Then
              lstrOutpuString &= String.Format("[{0}], ", lstrResultColumn)
            Else
              lstrOutpuString &= String.Format("{0}, ", lstrResultColumn)
            End If
          End If
        Next

        lstrOutpuString = lstrOutpuString.TrimEnd(",", " ")

        Return lstrOutpuString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SetCriteriaFromMetaHolder(ByVal lpMetaHolder As Core.IMetaHolder,
                                   Optional ByVal lpVersionIndex As Integer = 0,
                                   Optional ByRef lpErrorMessage As String = "")

      Dim lobjProperty As Core.ECMProperty = Nothing
      Dim lobjPropertyValue As String = ""
      Dim lobjDocument As Core.Document

      Try


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        For Each lobjCriterion As Criterion In Criteria

          Select Case lpMetaHolder.GetType.Name
            Case "Document"
              lobjDocument = lpMetaHolder
              If lobjDocument.PropertyExists(Core.PropertyScope.DocumentProperty, lobjCriterion.PropertyName) Then
                lobjProperty = lobjDocument.Properties(lobjCriterion.PropertyName)
              ElseIf lobjDocument.PropertyExists(Core.PropertyScope.VersionProperty, lobjCriterion.PropertyName) Then
                lobjProperty = lobjDocument.Versions(lpVersionIndex).Properties(lobjCriterion.PropertyName)
              ElseIf lobjDocument.PropertyExists(Core.PropertyScope.ContentProperty, lobjCriterion.Name) Then
                lobjProperty = lobjDocument.GetProperty(lobjCriterion.Name, Core.PropertyScope.ContentProperty, lpVersionIndex)
              End If

            Case Else
              If lpMetaHolder.Metadata.PropertyExists(lobjCriterion.PropertyName, False) Then
                lobjProperty = lpMetaHolder.Metadata.Item(lobjCriterion.PropertyName)
              End If
          End Select

          If lobjProperty Is Nothing Then
            lpErrorMessage = String.Format("Unable to get {0} property '" & lobjCriterion.PropertyName & "' from Criterion.", lpMetaHolder.GetType.Name)
            Exit Sub
          End If

          lobjCriterion.Cardinality = lobjProperty.Cardinality

          'Select Case lobjCriterion.PropertyScope
          '  Case Core.PropertyScope.DocumentProperty
          '    lobjProperty = lpDocument.Properties(lobjCriterion.PropertyName)

          '  Case Core.PropertyScope.VersionProperty
          '    lobjProperty = lpDocument.Versions(lpVersionIndex).Properties(lobjCriterion.PropertyName)

          'End Select

          lobjCriterion.DataType = lobjProperty.Type

          Select Case lobjProperty.Cardinality
            Case Core.Cardinality.ecmSingleValued
              lobjPropertyValue = lobjProperty.Value
              If lobjPropertyValue.Length > 0 Then
                lobjCriterion.Value = lobjProperty.Value
              End If

            Case Core.Cardinality.ecmMultiValued
              lobjCriterion.Operator = Criterion.pmoOperator.opIn

              ' Clear the values in case there are any remaining from a previous operation
              lobjCriterion.Values.Clear()

              For Each lstrValue As String In lobjProperty.Values
                If lstrValue.Length > 0 Then
                  lobjCriterion.Values.Add(lstrValue)
                End If
              Next

              ResultColumns.Add(lobjCriterion.Name)

          End Select

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::SetCriteriaFromMetaHolder(lpDocument:{0}, lpVersionIndex:{1}, lpErrorMessage:{2}", lpMetaHolder.Identifier, lpVersionIndex, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Exit Sub
      End Try

    End Sub
    Public Overridable Function BuildSQLString(Optional ByRef lpErrorMessage As String = "") As String
      Try
        Return BuildSQLString(False, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function BuildSQLString(lpGetCountOnly As Boolean, Optional ByRef lpErrorMessage As String = "") As String
      Dim lobjSQLBuilder As New StringBuilder
      Dim lstrTempSQL As String = String.Empty

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        ' First check to see if the Search.Criteria is in the first clause

        lobjSQLBuilder.Append("SELECT ")
        If (Me.DistinctQuery = True) AndAlso (lpGetCountOnly = False) Then
          lobjSQLBuilder.Append(" DISTINCT")
        End If

        If (Me.mIntLimitResults > 0) AndAlso (lpGetCountOnly = False) Then
          lobjSQLBuilder.AppendFormat(" TOP {0} ", Me.mIntLimitResults.ToString())
        End If

        If QueryTarget.Contains("."c) Then
          Dim lstrQueryTargetParts As String() = QueryTarget.Split(".")

          'lstrSQL &= ResultColumnsString & " FROM " ' & QueryTarget & "] WHERE"
          If lpGetCountOnly Then
            If QueryTarget.Contains(" "c) Then
              lobjSQLBuilder.AppendFormat("Count(*) FROM [{0}] WHERE", QueryTarget)
            Else
              lobjSQLBuilder.AppendFormat("Count(*) FROM {0} WHERE", QueryTarget)
            End If
          Else
            lobjSQLBuilder.AppendFormat("{0} FROM ", ResultColumnsString)
            For lintPartCounter As Integer = 0 To lstrQueryTargetParts.Length - 1
              If lstrQueryTargetParts(lintPartCounter).Contains(" "c) Then
                lobjSQLBuilder.AppendFormat("[{0}].", lstrQueryTargetParts(lintPartCounter))
              Else
                lobjSQLBuilder.AppendFormat("{0}.", lstrQueryTargetParts(lintPartCounter))
              End If
            Next
            lobjSQLBuilder.Remove(lobjSQLBuilder.Length - 1, 1)
            lobjSQLBuilder.Append(" WHERE")
          End If

          'For lintPartCounter As Integer = 0 To lstrQueryTargetParts.Length - 1
          '  lobjSQLBuilder.AppendFormat("[{0}].", lstrQueryTargetParts(lintPartCounter))
          'Next
          'lobjSQLBuilder.Remove(lobjSQLBuilder.Length - 1, 1)
          'lobjSQLBuilder.Append(" WHERE")
        Else
          If lpGetCountOnly Then
            If QueryTarget.Contains(" "c) Then
              lobjSQLBuilder.AppendFormat("Count(*) FROM [{1}] WHERE", ResultColumnsString, QueryTarget)
            Else
              lobjSQLBuilder.AppendFormat("Count(*) FROM {1} WHERE", ResultColumnsString, QueryTarget)
            End If
          Else
            If QueryTarget.Contains(" "c) Then
              lobjSQLBuilder.AppendFormat("{0} FROM [{1}] WHERE", ResultColumnsString, QueryTarget)
            Else
              lobjSQLBuilder.AppendFormat("{0} FROM {1} WHERE", ResultColumnsString, QueryTarget)
            End If
          End If
        End If


        Dim lstrEval As String = String.Empty
        Dim i As Integer = 0

        For Each cls As Clause In Me.Clauses

          If (cls.Criteria.Count = 0) Then
            Continue For
          End If

          If (i > 0) Then
            If (cls.SetEvaluation = Documents.Search.SetEvaluation.seAnd) Then
              lobjSQLBuilder.Append(" AND ")
            ElseIf (cls.SetEvaluation = Documents.Search.SetEvaluation.seOr) Then
              lobjSQLBuilder.Append(" OR ")
            End If
          End If

          lobjSQLBuilder.Append(" (")
          For Each lobjCriterion As Criterion In cls.Criteria



            If (lobjCriterion.SetEvaluation = Documents.Search.SetEvaluation.seAnd) Then
              lstrEval = " AND "
            ElseIf (lobjCriterion.SetEvaluation = Documents.Search.SetEvaluation.seOr) Then
              lstrEval = " OR "
            End If

            Select Case lobjCriterion.Cardinality
              Case Core.Cardinality.ecmSingleValued
                If (lobjCriterion.Value IsNot Nothing) AndAlso (lobjCriterion.Value.Length > 0) Then
                  ' Determine the data type of the criterion
                  Select Case lobjCriterion.DataType
                    Case Data.Criterion.pmoDataType.ecmString, Data.Criterion.pmoDataType.ecmGuid
                      Select Case lobjCriterion.Operator
                        Case Data.Criterion.pmoOperator.opLike, Data.Criterion.pmoOperator.opNotLike
                          lobjSQLBuilder.AppendFormat("{0}{1}%' " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"))
                        Case Data.Criterion.pmoOperator.opEndsWith
                          lobjSQLBuilder.AppendFormat("{0}{1}' " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"))
                        Case Data.Criterion.pmoOperator.opStartsWith
                          lobjSQLBuilder.AppendFormat("{0}{1}%' " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"))
                        Case Data.Criterion.pmoOperator.opContentContains
                          lobjSQLBuilder.AppendFormat("{0}{1}') " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"))
                          lobjSQLBuilder.Replace("WHERE", " d INNER JOIN VerityContentSearch cs ON cs.QueriedObject = d.This WHERE")
                        Case Else
                          lobjSQLBuilder.AppendFormat("{0}'{1}'" & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"))
                      End Select
                    Case Criterion.pmoDataType.ecmObject
                      Select Case lobjCriterion.Operator
                        Case Data.Criterion.pmoOperator.opLike, Data.Criterion.pmoOperator.opNotLike
                          lobjSQLBuilder.AppendFormat("{0}Object({1}%') " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "'')"))
                        Case Else
                          lobjSQLBuilder.AppendFormat("{0}Object('{1}')" & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "'')"))
                      End Select
                    Case Else
                      lobjSQLBuilder.AppendFormat("{0}{1}{2}", lobjCriterion.WhereClause, lobjCriterion.Value.Replace("'", "''"), lstrEval)
                  End Select

                Else
                  If lobjCriterion.Name.Contains(" "c) Then
                    lobjSQLBuilder.AppendFormat(" [{0}] IS NULL {1}", lobjCriterion.Name, lstrEval)
                  Else
                    lobjSQLBuilder.AppendFormat(" {0} IS NULL {1}", lobjCriterion.Name, lstrEval)
                  End If
                End If

              Case Core.Cardinality.ecmMultiValued

                If (lobjCriterion.Values.Count > 0) Then

                  Dim lstrSeparator As String = String.Empty
                  If (lobjCriterion.DataType = Criterion.pmoDataType.ecmString Or lobjCriterion.DataType = Criterion.pmoDataType.ecmDate) Then
                    lstrSeparator = "'"
                  End If
                  'lstrSQL &= "("
                  lobjSQLBuilder.Append("("c)
                  'lstrSQL &= lobjCriterion.Name & " in ("
                  lobjSQLBuilder.AppendFormat("{0} IN (", lobjCriterion.Name)
                  For Each lstrValue As String In lobjCriterion.Values
                    If lstrValue.Length > 0 Then
                      'lstrSQL &= "'" & lstrValue.Replace("'", "''") & "' in " & "[" & lobjCriterion.Name & "]" & lstrEval
                      lobjSQLBuilder.AppendFormat("'{0}' IN [{1}]{2},", lstrValue.Replace("'", "''"), lstrValue)
                      'lstrSQL &= lstrSeparator & lstrValue & lstrSeparator & ","
                      lobjSQLBuilder.AppendFormat("{0}{1}{0},", lstrSeparator, lstrValue)
                    End If
                  Next
                  If (lobjSQLBuilder.ToString.EndsWith(","c)) Then
                    lobjSQLBuilder.Remove(lobjSQLBuilder.ToString.LastIndexOf(","c), 1)
                  End If
                  lobjSQLBuilder.AppendFormat("){0}", lstrEval)
                End If
            End Select


          Next

          If lobjSQLBuilder.ToString.EndsWith(" AND ") Then
            lobjSQLBuilder.Remove(lobjSQLBuilder.ToString.LastIndexOf(" AND "), 4)
          ElseIf lobjSQLBuilder.ToString.EndsWith(" OR ") Then
            lobjSQLBuilder.Remove(lobjSQLBuilder.ToString.LastIndexOf(" OR "), 3)
          End If


          ' We cant do a trim end with StringBuilder but we can do this...
          lobjSQLBuilder = New StringBuilder(lobjSQLBuilder.ToString.TrimEnd)

          lobjSQLBuilder.Append(")"c)
          i += 1
        Next

        'If no criteria was defined you get this
        lobjSQLBuilder.Replace("WHERE ( [] IS NULL  )", String.Empty)
        lobjSQLBuilder.Replace("WHERE ( [] IS NULL)", String.Empty)
        lobjSQLBuilder.Replace("WHERE ([] IS NULL)", String.Empty)
        lobjSQLBuilder.Replace("WHERE ( IS NULL)", String.Empty)

        If lobjSQLBuilder.ToString.EndsWith(" WHERE") Then
          lobjSQLBuilder.Replace(" WHERE", String.Empty)
        End If

        If lobjSQLBuilder.ToString.EndsWith(" AND ") Then
          lobjSQLBuilder.Remove(lobjSQLBuilder.ToString.LastIndexOf(" AND"), 4)
        ElseIf lobjSQLBuilder.ToString.EndsWith(" OR ") Then
          lobjSQLBuilder.Remove(lobjSQLBuilder.ToString.LastIndexOf(" OR"), 3)
        End If


        If (mobjOrderBy.Count > 0) Then
          lobjSQLBuilder.Append(" ORDER BY ")
          For Each lobjOrderItem As OrderItem In mobjOrderBy
            If (lobjOrderItem.SortDirection = SortDirection.Asc) Then
              If lobjOrderItem.FieldName.Contains(" "c) Then
                lobjSQLBuilder.AppendFormat(" [{0}] ASC", lobjOrderItem.FieldName)
              Else
                lobjSQLBuilder.AppendFormat(" {0} ASC", lobjOrderItem.FieldName)
              End If
            Else
              If lobjOrderItem.FieldName.Contains(" "c) Then
                lobjSQLBuilder.AppendFormat(" [{0}] DESC", lobjOrderItem.FieldName)
              Else
                lobjSQLBuilder.AppendFormat(" {0} DESC", lobjOrderItem.FieldName)
              End If
            End If
            lobjSQLBuilder.Append(","c)
          Next
          lobjSQLBuilder = New StringBuilder(lobjSQLBuilder.ToString.TrimEnd(", "))
        End If

        lobjSQLBuilder = New StringBuilder(lobjSQLBuilder.ToString.TrimEnd(" "))

        Return lobjSQLBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "DataSource::BuildSQLString()")
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return lobjSQLBuilder.ToString
      End Try

    End Function


    Public Function BuildSQLStringORIGINAL(Optional ByRef lpErrorMessage As String = "") As String
      Dim lstrSQL As String = ""


      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        'For Each cls As Clause In Me.Clauses
        '  For Each cr As Criterion In cls.Criteria

        '  Next
        'Next

        lstrSQL &= "SELECT "

        If (Me.DistinctQuery = True) Then
          lstrSQL &= " DISTINCT"
        End If

        If (Me.mIntLimitResults > 0) Then
          lstrSQL &= " TOP " & Me.mIntLimitResults.ToString() & " "
        End If
        lstrSQL &= ResultColumnsString & " FROM [" & QueryTarget & "] WHERE"


        For Each lobjCriterion As Data.Criterion In Criteria
          Dim lstrEval As String = ""
          If (lobjCriterion.SetEvaluation = Documents.Search.SetEvaluation.seAnd) Then
            lstrEval = " AND "
          ElseIf (lobjCriterion.SetEvaluation = Documents.Search.SetEvaluation.seOr) Then
            lstrEval = " OR "
          End If

          Select Case lobjCriterion.Cardinality
            Case Core.Cardinality.ecmSingleValued
              If (lobjCriterion.Value IsNot Nothing) AndAlso (lobjCriterion.Value.Length > 0) Then
                ' Determine the data type of the criterion
                Select Case lobjCriterion.DataType
                  Case Data.Criterion.pmoDataType.ecmString, Data.Criterion.pmoDataType.ecmGuid
                    Select Case lobjCriterion.Operator
                      Case Data.Criterion.pmoOperator.opLike, Data.Criterion.pmoOperator.opNotLike
                        lstrSQL &= String.Format("{0}{1}%' " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value)
                      Case Else
                        lstrSQL &= String.Format("{0}'{1}' " & lstrEval, lobjCriterion.WhereClause, lobjCriterion.Value)
                    End Select
                  Case Else
                    lstrSQL &= lobjCriterion.WhereClause & lobjCriterion.Value & lstrEval
                End Select

              Else
                lstrSQL &= " [" & lobjCriterion.Name & "] IS NULL " & lstrEval
              End If

            Case Core.Cardinality.ecmMultiValued

              If (lobjCriterion.Values.Count > 0) Then



                lstrSQL &= "("
                For Each lstrValue As String In lobjCriterion.Values
                  If lstrValue.Length > 0 Then
                    lstrSQL &= "'" & lstrValue & "' in " & "[" & lobjCriterion.Name & "]" & lstrEval
                  End If
                Next
                If lstrSQL.EndsWith(" AND ") Then
                  lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" AND"), 4)
                ElseIf lstrSQL.EndsWith(" OR ") Then
                  lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" OR"), 3)
                End If
                lstrSQL &= ")" & lstrEval
              End If
          End Select

        Next

        If lstrSQL.EndsWith(" AND ") Then
          lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" AND"), 4)
        ElseIf lstrSQL.EndsWith(" OR ") Then
          lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" OR"), 3)
        End If


        If (mobjOrderBy.Count > 0) Then
          lstrSQL &= " ORDER BY "
          For Each lobjOrderItem As OrderItem In mobjOrderBy
            Dim sortdir As String = "DESC"
            If (lobjOrderItem.SortDirection = SortDirection.Asc) Then
              sortdir = "ASC"
            End If
            lstrSQL &= " [" & lobjOrderItem.FieldName & "] " & sortdir & ","
          Next
          lstrSQL = lstrSQL.TrimEnd(",")
        End If

        Return lstrSQL

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "DataSource::BuildSQLString()")
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return lstrSQL
      End Try



      'Dim lstrSQL As String = ""

      'Try

      '          lstrSQL &= "SELECT "

      '          If (Me.mIntLimitResults > 0) Then
      '              lstrSQL &= " TOP " & Me.mIntLimitResults.ToString()
      '          End If
      '          lstrSQL &= ResultColumnsString & " FROM [" & QueryTarget & "] WHERE"

      '          For Each lobjCriterion As Criterion In Criteria
      '              '  Select Case lobjCriterion.PropertyScope
      '              '    Case Core.PropertyScope.DocumentProperty
      '              '      lobjProperty = lpDocument.Properties(lobjCriterion.PropertyName)

      '              '    Case Core.PropertyScope.VersionProperty
      '              '      lobjProperty = lpDocument.Versions(lpVersionIndex).Properties(lobjCriterion.PropertyName)

      '              '  End Select

      '              Select Case lobjCriterion.Cardinality
      '                  Case Core.Cardinality.ecmSingleValued
      '                      If (lobjCriterion.Value IsNot Nothing) AndAlso (lobjCriterion.Value.Length > 0) Then
      '                          lstrSQL &= lobjCriterion.WhereClause & "'" & lobjCriterion.Value & "' AND"
      '                      Else
      '                          lstrSQL &= " [" & lobjCriterion.Name & "] IS NULL AND"
      '                      End If

      '                  Case Core.Cardinality.ecmMultiValued
      '                      lobjCriterion.Operator = Criterion.pmoOperator.opIn

      '                      lstrSQL &= lobjCriterion.WhereClause

      '                      For Each lstrValue As String In lobjCriterion.Values
      '                          If lstrValue.Length > 0 Then
      '                              lstrSQL &= "'" & lstrValue & "', "
      '                          End If
      '                      Next
      '                      lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(", "), 2)
      '                      lstrSQL &= ") AND"

      '              End Select

      '          Next

      '          If lstrSQL.EndsWith(" AND") Then
      '              lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" AND"), 4)
      '          ElseIf lstrSQL.EndsWith(" OR") Then
      '              lstrSQL = lstrSQL.Remove(lstrSQL.LastIndexOf(" OR"), 3)
      '          End If


      '          If (mstrOrderBy.Length > 0) Then
      '              lstrSQL &= " ORDER BY " & mstrOrderBy.Trim()
      '          End If

      '          Return lstrSQL

      '      Catch ex As Exception
      '          ApplicationLogging.LogException(ex, "DataSource::BuildSQLString()")
      '          lpErrorMessage = FormatCallStack(ex)
      '          Return lstrSQL
      '      End Try

    End Function

    Public Overridable Function BuildSQLString(ByVal lpMetaHolder As Core.IMetaHolder,
                                               Optional ByVal lpVersionIndex As Integer = 0,
                                               Optional ByRef lpErrorMessage As String = "") As String
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        SetCriteriaFromMetaHolder(lpMetaHolder, lpVersionIndex, lpErrorMessage)

        If lpErrorMessage.Length > 0 Then
          Return String.Empty
        End If

        Return BuildSQLString(lpErrorMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "DataSource::BuildSQLString(lpVersion)")
        Return ""
      End Try
    End Function

    Private Sub InitializeConnection()
      Try
        mobjConnection = New OleDb.OleDbConnection(ConnectionString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "DataSource::InitializeConnection")
      End Try
    End Sub

    Public Shared Function GetInfoFromString(ByVal lpInputString As String,
                                         ByVal lpAttribute As String,
                                         Optional ByVal lpDelimiter As String = ";") As String

      Try
        Dim lstrNameValuePairs() As String = lpInputString.Split(lpDelimiter)
        Dim lstrNameValuePair() As String
        Dim lblnInfoFound As Boolean = False

        For Each lstrPair As String In lstrNameValuePairs
          'debug.writeline(lstrPair)
          lstrNameValuePair = lstrPair.Split("=")
          If lstrNameValuePair(0) = lpAttribute Then
            lblnInfoFound = True
            Return lstrNameValuePair(1)
          End If
        Next

        If lblnInfoFound = False Then
          Throw New Exception(String.Format("'{0}' not found in '{1}'", lpAttribute, lpInputString))
        End If

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::GetInfoFromString(lpInputString:'{0}', lpAttribute:'{1}', lpDelimiter:'{2}')",
                                      lpInputString, lpAttribute, lpDelimiter))
        Return Nothing
      End Try

    End Function


#End Region

#Region "DataLookup Implementation"

    Private Function RetrieveValue(ByVal lpSQL As String) As Object

      Try
        Dim lobjValue As Object = Nothing

        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        ' Create a query command on the connection
        Dim lobjCommand As New OleDb.OleDbCommand(lpSQL, Connection)

        ' Run the query; get the DataReader object
        Dim lobjDataReader As OleDb.OleDbDataReader

        Try
          lobjDataReader = lobjCommand.ExecuteReader(CommandBehavior.SingleRow)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("DataSource::RetrieveValue_ExecuteReader('{0}')", lpSQL))
          Throw New Exception("Unable to get value [" & ex.Message & "]", ex)
        End Try

        If lobjDataReader.HasRows Then
          ' Read the resultset
          Do While lobjDataReader.Read
            ' Get the first value
            lobjValue = lobjDataReader.Item(Me.SourceColumn)
            Exit Do
          Loop

        Else
          'Throw New Exception("No value found for the expression (" & lpSQL & ")")
          ApplicationLogging.WriteLogEntry(New ValueNotFoundException(lpSQL).Message, TraceEventType.Warning, 44404)
          Return Nothing
        End If

        ' Always close the DataReader
        lobjDataReader.Close()

        ' Always close the Connection
        If Connection.State = ConnectionState.Open Then
          Connection.Close()
        End If

        Return lobjValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::RetrieveValue('{0}')", lpSQL))
        Return Nothing
      End Try

    End Function

    Public Overrides Function GetValue(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        '  Make sure that we do not have any empty clauses
        ClearEmptyClauses()

        Return RetrieveValue(SQLStatement(lpMetaHolder))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::GetValue(lpMetaHolder:'{0}')", lpMetaHolder.Identifier))
        Return Nothing
      End Try
    End Function

    Public Overrides Function SourceExists(ByVal lpMetaHolder As Core.IMetaHolder) As Boolean
      Try
        If Connection IsNot Nothing Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function RetrieveValues(ByVal lpSQL As String) As Object

      Try
        Dim lobjValue As New MapList(New DataList())

        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        ' Create a query command on the connection
        Dim lobjCommand As New OleDb.OleDbCommand(lpSQL, Connection)

        ' Run the query; get the DataReader object
        Dim lobjDataReader As OleDb.OleDbDataReader

        Try
          lobjDataReader = lobjCommand.ExecuteReader(CommandBehavior.Default)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("DataSource::RetrieveValue_ExecuteReader('{0}')", lpSQL))
          Throw New Exception("Unable to get value [" & ex.Message & "]", ex)
        End Try

        If lobjDataReader.HasRows Then
          ' Read the resultset
          Do While lobjDataReader.Read
            'lobjValue = lobjDataReader.Item(Me.SourceColumn)
            lobjValue.Add(lobjDataReader.Item(Me.ResultColumns(0)), lobjDataReader.Item(Me.SourceColumn))
          Loop

        Else
          'Throw New Exception("No value found for the expression (" & lpSQL & ")")
          ApplicationLogging.WriteLogEntry(New ValueNotFoundException(lpSQL).Message.Replace("value found", "values found"),
                                           TraceEventType.Warning, 45404)
          Return Nothing
        End If

        ' Always close the DataReader
        lobjDataReader.Close()

        ' Always close the Connection
        If Connection.State = ConnectionState.Open Then
          Connection.Close()
        End If

        Return lobjValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::RetrieveValue('{0}')", lpSQL))
        Return Nothing
      End Try

    End Function

    Public Overrides Function GetValues(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Try

        Dim lobjMapList As MapList

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        If Connection.State = ConnectionState.Closed Then
          Connection.Open()
        End If

        '  Make sure that we do not have any empty clauses
        ClearEmptyClauses()

        ResultColumns.Clear()
        ResultColumns.Add(SourceColumn)

        lobjMapList = RetrieveValues(SQLStatement(lpMetaHolder))

        'If SourceColumn.ToLower.EndsWith("_mv") Then
        '  ' This is supposed to be a multi-valued property
        '  ' Try to split it.
        '  Dim lstrSplitValues() As String
        '  If Delimiter IsNot Nothing AndAlso Delimiter.Length > 0 Then
        '    Dim lobjSplitMapList As New MapList(New DataList())
        '    For Each lobjValueMap As ValueMap In lobjMapList
        '      lstrSplitValues = lobjValueMap.Original.Split(Delimiter)
        '      For lintValueCounter As Integer = 0 To lstrSplitValues.Length - 1
        '        lobjSplitMapList.Add(lstrSplitValues(lintValueCounter), lobjValueMap.Replacement)
        '      Next
        '    Next
        '    Return lobjSplitMapList
        '  End If
        'End If

        Return lobjMapList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DataSource::GetValues(lpMetaHolder:'{0}')", lpMetaHolder.Identifier))
        Return Nothing
      End Try
    End Function

    Private Sub ClearEmptyClauses()
      Try
        If Me.Clauses Is Nothing Then
          Exit Sub
        End If

        Dim lobjTestClause As Clause

        For lintClauseCounter As Integer = Me.Clauses.Count - 1 To 0 Step -1
          lobjTestClause = Me.Clauses(lintClauseCounter)
          If lobjTestClause.Criteria Is Nothing OrElse lobjTestClause.Criteria.Count = 0 Then
            Me.Clauses.Remove(lobjTestClause)
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IPropertyLookup Implementation"

    Public Overridable Property SourceProperty() As LookupProperty Implements IPropertyLookup.SourceProperty
      Get
        Try
          'Throw New NotImplementedException
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As LookupProperty)
        Try
          'Throw New NotImplementedException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable Property DestinationProperty() As LookupProperty Implements IPropertyLookup.DestinationProperty
      Get
        Try
          ' Throw New NotImplementedException
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As LookupProperty)
        Try
          'Throw New NotImplementedException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property ObjSearchType() As SearchType
      Get
        Return mobjSearchType
      End Get
    End Property


#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return DATA_SOURCE_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.DeSerialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Serialize(lpFilePath, "")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        Serializer.Serialize.XmlFile(Me, lpFilePath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


      Return Serializer.Serialize.Xml(Me)

    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If


      Return Serializer.Serialize.XmlString(Me)

    End Function

    Public Class InvalidConnectionStringException
      Inherits ApplicationException

      Public Sub New(ByVal Message As String)
        MyBase.New(Message)
      End Sub

      Public Sub New(ByVal Message As String, ByVal InnerException As Exception)
        MyBase.New(Message, InnerException)
      End Sub

    End Class

#End Region

    'Protected Overrides Sub Finalize()
    '  MyBase.Finalize()
    'End Sub

#Region "ILoggable Implementation"

    'Private mobjLogSession As Gurock.SmartInspect.Session

    'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
    '  Try
    '    ApplicationLogging.FinalizeLogSession(mobjLogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.MediumAquamarine)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
    '  Get
    '    Try
    '      If mobjLogSession Is Nothing Then
    '        InitializeLogSession()
    '      End If
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: free other state (managed objects).
            mobjConnection.Dispose()
          End If

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' DISPOSETODO: set large fields to null.
        End If
        Me.disposedValue = True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

    Private Sub MobjResultColumns_CollectionChanged(sender As Object, e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles MobjResultColumns.CollectionChanged

      Try
        If e.Action = Specialized.NotifyCollectionChangedAction.Remove Then
          For Each lstrColumn As String In e.OldItems
            If (mobjRemovedColumns.Contains(lstrColumn) = False) Then
              mobjRemovedColumns.Add(lstrColumn)
            End If
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try


    End Sub

  End Class

End Namespace