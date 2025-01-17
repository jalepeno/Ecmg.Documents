'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  DbHelper.vb
'   Description :  [type_description_here]
'   Created     :  4/25/2013 2:49:17 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports Microsoft.Data.SqlClient

#End Region

Namespace Utilities

  Public Class DbHelper

#Region "Static Variables"

    Private Shared mstrOleDbConnectionString As String = String.Empty

#End Region

#Region "Public Shared Methods"

    Public Shared Function GetAvailableProviderFactoryNames() As IList(Of String)
      Try
        Dim lobjList As New List(Of String)

        Dim lobjFactoryClasses As DataTable = DbProviderFactories.GetFactoryClasses

        If lobjFactoryClasses Is Nothing OrElse lobjFactoryClasses.Rows.Count = 0 Then
          Return lobjList
        End If

        For Each lobjDataRow As DataRow In lobjFactoryClasses.Rows
          lobjList.Add(lobjDataRow(0))
        Next

        lobjList.Sort()

        Return lobjList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     Gets a dictionary containing the description and invariant name of the available .Net database provider factories on the current machine.
    ''' </summary>
    ''' <returns>
    '''     An IDictionary(Description, InvariantName)
    ''' </returns>
    Public Shared Function GetAvailableProviderFactoryDictionary() As IDictionary(Of String, String)
      Try

        Dim lobjDictionary As New SortedDictionary(Of String, String)

        Dim lobjFactoryClasses As DataTable = DbProviderFactories.GetFactoryClasses

        If lobjFactoryClasses Is Nothing OrElse lobjFactoryClasses.Rows.Count = 0 Then
          Return lobjDictionary
        End If

        For Each lobjDataRow As DataRow In lobjFactoryClasses.Rows
          lobjDictionary.Add(lobjDataRow(1), lobjDataRow(2))
        Next

        Return lobjDictionary

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function GetDbProviderFactory() As DbProviderFactory
    '  Try
    '    Dim lobjProviderFactory As DbProviderFactory = DbProviderFactories.GetFactory(lstrInvariantName)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Shared Function BuildSQLConnectionString(lpDataSource As String) As String
      Try
        Dim lobjBuilder As New SqlConnectionStringBuilder

        With lobjBuilder
          .DataSource = lpDataSource
          .IntegratedSecurity = True
        End With

        Return lobjBuilder.ConnectionString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function BuildSQLConnectionString(lpDataSource As String, lpInitialCatalog As String) As String
      Try
        Dim lobjBuilder As New SqlConnectionStringBuilder

        With lobjBuilder
          .DataSource = lpDataSource
          .IntegratedSecurity = True
          .InitialCatalog = lpInitialCatalog
        End With

        Return lobjBuilder.ConnectionString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function BuildSQLConnectionString(lpDataSource As String, lpUserName As String, lpPassword As String) As String
      Try
        Dim lobjBuilder As New SqlConnectionStringBuilder

        With lobjBuilder
          .DataSource = lpDataSource
          .UserID = lpUserName
          .Password = lpPassword
        End With

        Return lobjBuilder.ConnectionString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function BuildSQLConnectionString(lpDataSource As String, lpUserName As String, lpPassword As String, lpInitialCatalog As String, lpApplicationName As String) As String
      Try
        Dim lobjBuilder As New SqlConnectionStringBuilder

        With lobjBuilder
          .DataSource = lpDataSource
          If Not String.IsNullOrEmpty(lpUserName) Then
            .UserID = lpUserName
          End If
          If Not String.IsNullOrEmpty(lpPassword) Then
            .Password = lpPassword
          End If
          If (String.IsNullOrEmpty(lpUserName) AndAlso String.IsNullOrEmpty(lpPassword)) Then
            .IntegratedSecurity = True
          End If
          If Not String.IsNullOrEmpty(lpInitialCatalog) Then
            .InitialCatalog = lpInitialCatalog
          End If
          If Not String.IsNullOrEmpty(lpApplicationName) Then
            .ApplicationName = lpApplicationName
          End If
        End With

        Return lobjBuilder.ConnectionString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetDatabaseSchemaInfo(lpConnectionString As String) As DataSet
      Try

        Dim lobjSchemaSet As New DataSet("DatabaseSchemas")
        Dim lobjCatalogTable As DataTable
        Dim lobjTableInfo As DataTable
        Dim lobjCompleteTableInfo As DataTable = Nothing
        Dim lobjCatalogs As New List(Of String)
        Dim lobjTables As New List(Of String)
        Dim lstrCatalogConnectionString As String = Nothing
        Dim lintCatalogCounter As Integer = 0

        mstrOleDbConnectionString = lpConnectionString

        'Dim lobjConnection As New System.Data.OleDb.OleDbConnection(lpOleDbConnectionString) '.SqlClient.SqlConnection(lstrConnectionString)
        Dim lobjConnection As Common.DbConnection = DbHelper.CreateConnection(lpConnectionString)

        lobjConnection.Open()

        'lobjCatalogTable = lobjSchemaSet.Tables.Add("Catalogs")

        lobjCatalogTable = lobjConnection.GetSchema("Catalogs")

        lobjSchemaSet.Tables.Add(lobjCatalogTable)

        For Each lobjRow As DataRow In lobjCatalogTable.Rows
          lobjCatalogs.Add(lobjRow.Item(0))
        Next

        lobjConnection.Close()

        For Each lstrCatalog As String In lobjCatalogs

          lstrCatalogConnectionString = String.Format("{0};Database={1}", lpConnectionString, lstrCatalog)
          lobjConnection.ConnectionString = lstrCatalogConnectionString
          lobjConnection.Open()
          lobjTableInfo = lobjConnection.GetSchema("Tables")
          lobjConnection.Close()
          'lobjTables = New List(Of String)

          If lintCatalogCounter = 0 Then
            lobjCompleteTableInfo = lobjTableInfo.Clone
          End If

          For Each lobjRow As DataRow In lobjTableInfo.Rows
            'lobjCompleteTableInfo.Rows.Add(lobjRow)
            lobjCompleteTableInfo.ImportRow(lobjRow)
          Next

          lintCatalogCounter += 1
        Next

        lobjSchemaSet.Tables.Add(lobjCompleteTableInfo)

        Return lobjSchemaSet

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateConnection(lpConnectionString As String) As IDbConnection
      Dim lobjConnection As IDbConnection
      Try

        If String.IsNullOrEmpty(lpConnectionString) Then
          Throw New ArgumentNullException("lpConnectionString")
        End If

        If lpConnectionString.ToLower.StartsWith("provider=") Then
          lobjConnection = New OleDbConnection(lpConnectionString)
        ElseIf lpConnectionString.ToLower.StartsWith("driver=") Then
          lobjConnection = New OdbcConnection(lpConnectionString)
        Else
          lobjConnection = New Microsoft.Data.SqlClient.SqlConnection(lpConnectionString)
        End If

        Return lobjConnection

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetDatabaseNames(lpConnectionString As String) As List(Of String)
      Dim lobjConnection As IDbConnection = Nothing
      Try

        lobjConnection = DbHelper.CreateConnection(lpConnectionString)

        lobjConnection.Open()
        lobjConnection.ChangeDatabase("master")

        Return GetDatabaseNames(lobjConnection)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjConnection IsNot Nothing AndAlso lobjConnection.State = ConnectionState.Open Then
          lobjConnection.Close()
        End If
      End Try
    End Function

    Public Shared Function GetDatabaseNames(lpConnection As IDbConnection) As List(Of String)
      Try

        Dim lobjReturnList As New List(Of String)

        Using lpConnection
          If lpConnection.State = ConnectionState.Closed Then
            lpConnection.Open()
          End If

          If Not String.Equals(lpConnection.Database, "master") Then
            lpConnection.ChangeDatabase("master")
          End If

          Dim lobjDatabaseCatalog As DataTable

          If TypeOf lpConnection Is OleDbConnection Then
            lobjDatabaseCatalog = DirectCast(lpConnection, OleDbConnection).GetSchema("Catalogs")
          ElseIf TypeOf lpConnection Is SqlConnection Then
            lobjDatabaseCatalog = DirectCast(lpConnection, SqlConnection).GetSchema("Databases")
          Else
            Throw New InvalidOperationException("Unable to get database names, connection is unknown type.")
          End If

          lpConnection.Close()

          For Each lobjRow As DataRow In lobjDatabaseCatalog.Rows
            lobjReturnList.Add(lobjRow.Item(0))
          Next

          lobjReturnList.Sort()

          Return lobjReturnList

        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetEmptyDatabaseNames(lpConnectionString As String) As List(Of String)

      Dim lobjConnection As IDbConnection = Nothing

      Try

        mstrOleDbConnectionString = lpConnectionString

        If String.IsNullOrEmpty(lpConnectionString) Then
          Throw New ArgumentNullException("lpOleDbConnectionString")
        End If

        lobjConnection = DbHelper.CreateConnection(lpConnectionString)

        Return GetEmptyDatabaseNames(lobjConnection)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjConnection IsNot Nothing AndAlso lobjConnection.State = ConnectionState.Open Then
          lobjConnection.Close()
        End If
      End Try
    End Function

    Public Shared Function GetEmptyDatabaseNames(lpConnection As IDbConnection) As List(Of String)
      Try

        Dim lobjDatabaseNames As List(Of String) = GetDatabaseNames(lpConnection)
        Dim lobjEmptyDatabases As New List(Of String)
        Dim lintBaseTableCount As Integer

        For Each lstrDatabase As String In lobjDatabaseNames
          ' Skip the model database
          If String.Equals(lstrDatabase, "model") Then
            Continue For
          End If

          ' For some reason we keep losing the connection string
          If String.IsNullOrEmpty(lpConnection.ConnectionString) Then
            lpConnection.ConnectionString = mstrOleDbConnectionString
          End If

          Try
            ' We will put this inside a special try since we might not be able to open all databases.
            lintBaseTableCount = GetBaseTableCount(lpConnection, lstrDatabase)
            If lintBaseTableCount = 0 Then
              lobjEmptyDatabases.Add(lstrDatabase)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            Continue For
          End Try
        Next

        Return lobjEmptyDatabases

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetBaseTableCount(lpConnectionString As String) As Integer
      Dim lobjConnection As IDbConnection = Nothing
      Try
        lobjConnection = DbHelper.CreateConnection(lpConnectionString)

        Return GetBaseTableCount(lobjConnection, lobjConnection.Database)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjConnection IsNot Nothing AndAlso lobjConnection.State = ConnectionState.Open Then
          lobjConnection.Close()
        End If
      End Try
    End Function

    Private Shared Function GetBaseTableCount(lpConnection As IDbConnection, lpDatabaseName As String) As Integer
      Try

        Dim lstrSQL As String = "SELECT COUNT(*) from information_schema.tables WHERE table_type = 'base table'"
        Dim lintTableCount As Integer = 0

        Using lpConnection
          If lpConnection.State = ConnectionState.Closed Then
            lpConnection.Open()
          End If

          lpConnection.ChangeDatabase(lpDatabaseName)

          Using lobjCommand As IDbCommand = lpConnection.CreateCommand
            lobjCommand.CommandText = lstrSQL
            lintTableCount = lobjCommand.ExecuteScalar()
            lobjCommand.Dispose()
          End Using
        End Using

        Return lintTableCount

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetTableNames(lpConnectionString As String) As List(Of String)
      Dim lobjConnection As IDbConnection = Nothing
      Try
        lobjConnection = DbHelper.CreateConnection(lpConnectionString)
        lobjConnection.Open()
        Return GetTableNames(lobjConnection)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjConnection IsNot Nothing AndAlso lobjConnection.State = ConnectionState.Open Then
          lobjConnection.Close()
        End If
      End Try
    End Function

    Public Shared Function GetTableNames(lpConnection As IDbConnection) As List(Of String)
      Try

        Dim lobjReturnList As New List(Of String)

        Dim lobjTableCatalog As DataTable = GetTableCatalog(lpConnection)

        For Each lobjRow As DataRow In lobjTableCatalog.Rows
          If ((lobjRow(3) = "TABLE") OrElse (lobjRow(3) = "BASE TABLE")) Then
            lobjReturnList.Add(lobjRow.Item(2))
          End If
        Next

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetDatabaseCatalog(lpConnectionString As String) As DataTable
      Try

        mstrOleDbConnectionString = lpConnectionString

        ' Dim lobjConnection As New System.Data.OleDb.OleDbConnection(lpOleDbConnectionString)
        Dim lobjConnection As IDbConnection = DbHelper.CreateConnection(lpConnectionString)
        lobjConnection.Open()
        lobjConnection.ChangeDatabase("master")

        Return GetDatabaseCatalog(lobjConnection)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Shared Function GetDatabaseCatalog(lpConnection As IDbConnection) As DataTable
      Try
        Using lpConnection
          If lpConnection.State = ConnectionState.Closed Then
            lpConnection.Open()
          End If

          If Not String.Equals(lpConnection.Database, "master") Then
            lpConnection.ChangeDatabase("master")
          End If

          Dim lobjDatabaseCatalog As DataTable = Nothing

          If TypeOf lpConnection Is OleDbConnection Then
            lobjDatabaseCatalog = DirectCast(lpConnection, OleDbConnection).GetSchema("Catalogs")
          ElseIf TypeOf lpConnection Is SqlConnection Then
            lobjDatabaseCatalog = DirectCast(lpConnection, SqlConnection).GetSchema("Databases")
          Else
            Throw New InvalidOperationException("Unable to get database names, connection is unknown type.")
          End If

          Return lobjDatabaseCatalog

        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetTableCatalog(lpConnection As IDbConnection) As DataTable
      Try

        Using lpConnection
          If lpConnection.State = ConnectionState.Closed Then
            lpConnection.Open()
          End If

          Dim lobjTableCatalog As DataTable

          If TypeOf lpConnection Is OleDbConnection Then
            lobjTableCatalog = DirectCast(lpConnection, OleDbConnection).GetSchema("Tables")
          ElseIf TypeOf lpConnection Is SqlConnection Then
            lobjTableCatalog = DirectCast(lpConnection, SqlConnection).GetSchema("Tables")
          Else
            Throw New InvalidOperationException("Unable to get table names, connection is unknown type.")
          End If

          Return lobjTableCatalog

        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
