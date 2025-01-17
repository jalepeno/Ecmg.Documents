'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Data
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Data
Imports Documents.Utilities

Namespace Providers
  ''' <summary>Base search class from which all provider specific search objects inherit.</summary>
  <Serializable()>
  Public MustInherit Class CSearch
    Implements IDisposable
    Implements ISearch

#Region "Class Constants"

    Public Const ID_COLUMN As String = "Id"
    Public Const DOCUMENT_QUERY_TARGET As String = "Document"

#End Region

#Region "Class Variables"

    Private mobjProvider As IProvider
    Private mobjCriteria As New Criteria
    Private mobjSearchResultSet As SearchResultSet
    Private mobjDataSource As New Data.DataSource(Me)
    Private mobjDefaultResultColumns As Data.ResultColumns

#End Region

#Region "Constructors"

    ''' <summary>Default Constructor</summary>
    Public Sub New()

    End Sub

    ''' <summary>Constructs a new Search object using the specified provider.</summary>
    Public Sub New(ByRef lpProvider As IProvider)
      Provider = lpProvider
    End Sub

    ''' <summary>
    ''' Constructs a new Search object using the specified provider, IDColumn and
    ''' QueryTarget values.
    ''' </summary>
    Public Sub New(ByRef lpProvider As IProvider,
                   ByVal lpIdColumn As String,
                   ByVal lpQueryTarget As String)

      With DataSource
        .SourceColumn = lpIdColumn
        .QueryTarget = lpQueryTarget
      End With
      Provider = lpProvider

    End Sub

    ''' <summary>
    ''' Constructs a new Search object using the specified provider, Criteria, IDColumn
    ''' and QueryTarget values.
    ''' </summary>
    Public Sub New(ByRef lpProvider As IProvider,
                   ByVal lpCriteria As Criteria,
                   ByVal lpIdColumn As String,
                   ByVal lpQueryTarget As String)

      Me.New(lpProvider, lpIdColumn, lpQueryTarget)
      Criteria = lpCriteria

    End Sub

    ''' <summary>
    ''' Constructs a new Search object using the specified provider, Criteria, IDColumn
    ''' , DataSource and QueryTarget values.
    ''' </summary>
    Public Sub New(ByRef lpProvider As IProvider,
                   ByVal lpCriteria As Criteria,
                   ByVal lpIdColumn As String,
                   ByVal lpQueryTarget As String, ByVal lpDataSource As Data.DataSource)

      Me.New(lpProvider, lpIdColumn, lpQueryTarget, lpDataSource)
      Criteria = lpCriteria

    End Sub

    ''' <summary>
    ''' Constructs a new Search object using the specified provider, IDColumn
    ''' , DataSource and QueryTarget values.
    ''' </summary>
    Public Sub New(ByRef lpProvider As IProvider,
                   ByVal lpIdColumn As String,
                   ByVal lpQueryTarget As String, ByVal lpDataSource As Data.DataSource)

      ' Set the data source
      ' This is primarily used in cases where 
      ' the provider needs to override the 
      ' data source with a new one inherited from the original

      mobjDataSource = lpDataSource
      With DataSource
        .SourceColumn = lpIdColumn
        .QueryTarget = lpQueryTarget
      End With
      Provider = lpProvider

    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' A collection of criterion objects. Used to specify the objects to search
    ''' for.
    ''' </summary>
    Public Property Criteria() As Criteria Implements ISearch.Criteria
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjCriteria

      End Get
      Set(ByVal value As Criteria)

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        mobjCriteria = value

      End Set
    End Property

    ''' <summary>
    ''' Defines the default query target for the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This is to be specified by the provider implementation</remarks>
    Public MustOverride ReadOnly Property DefaultQueryTarget As String Implements ISearch.DefaultQueryTarget

    Protected Overridable ReadOnly Property DefaultDelimitedResultColumns As String
      Get
        Return "Id,Title"
      End Get
    End Property

    ''' <summary>
    ''' Defines the default result columns for the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This is to be specified by the provider implementation</remarks>
    Public ReadOnly Property DefaultResultColumns As ResultColumns Implements ISearch.DefaultResultColumns
      Get
        Try
          If mobjDefaultResultColumns Is Nothing Then
            InitializeDefaultResultColumns()
          End If
          Return mobjDefaultResultColumns
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>An object containing the objects returned from the search.</summary>
    Public ReadOnly Property SearchResultSet() As SearchResultSet Implements ISearch.SearchResultSet
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjSearchResultSet

      End Get
    End Property

    ''' <summary>The DataSource object used to build the search.</summary>
    Public Overridable ReadOnly Property DataSource() As Data.DataSource Implements ISearch.DataSource
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjDataSource

      End Get
    End Property

    ''' <summary>The provider used to interact with the content repository.</summary>
    <Xml.Serialization.XmlIgnore()>
    Public Property Provider() As IProvider Implements ISearch.Provider
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjProvider

      End Get
      Set(ByVal value As IProvider)

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        mobjProvider = value

      End Set
    End Property

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Excutes the search using the search methodology of the content provider's
    ''' API.
    ''' </summary>
    Public MustOverride Function Execute(ByVal Args As SearchArgs) As SearchResultSet Implements ISearch.Execute

    Public MustOverride Function Execute() As SearchResultSet Implements ISearch.Execute

    'Public MustOverride Function SimpleSearch(ByVal Args As SimpleSearchArgs) As SearchResultSet Implements ISearch.SimpleSearch
    Public MustOverride Function SimpleSearch(ByVal Args As SimpleSearchArgs) As DataTable Implements ISearch.SimpleSearch

    Public Event SearchUpdate(ByVal sender As Object, ByVal e As SearchEventArgs) Implements ISearch.SearchUpdate

    Public Event SearchComplete(ByVal sender As Object, ByVal e As SearchEventArgs) Implements ISearch.SearchComplete

    ''' <summary>
    ''' Clears out all of the clauses, criteria and ordering
    ''' </summary>
    ''' <remarks>Call before defining a new search</remarks>
    Public Overridable Sub Clear()
      Try
        'Clear out the search
        Me.Criteria.Clear()
        Me.DataSource.Reset()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Reset Collections
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub Reset() Implements ISearch.Reset


      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Clear()

        'Me.DataSource.OrderBy.Clear()
        'Me.Criteria.Clear()
        'Me.DataSource.ResultColumns.Clear()
        'Me.DataSource.Clauses.Clear()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Protected Methods"
    ''' <summary>
    ''' Initializes the Search to be performed.
    ''' </summary>
    ''' <param name="Args">The SearchArgs passed into Search.Execute</param>
    ''' <remarks>Intended to be called from within Execute method.</remarks>
    Protected Overridable Sub InitializeSearch(ByVal Args As SearchArgs)

      Dim lstrErrorMessage As String = ""

      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        '  Make sure the Provider property is set.
        If Provider Is Nothing Then
          ApplicationLogging.WriteLogEntry(String.Format("Error in {0}::CSearch::InitializeSearch: 'The Provider property is not set, cannot initialize the search.'", Me.GetType.Name))
          Throw New InvalidOperationException("The Provider property is not set, cannot initialize the search")
        End If

        ' Initialize the DataSource with the correct QueryTarget and IDColumn
        InitializeDataSource(ID_COLUMN, DOCUMENT_QUERY_TARGET)

        ' Occasionally the Criteria for the DataSource is already set and Me.Criteria is not set
        ' In that case we do not want to zero out the criteria
        ' This is normally the case for the console Exporter app
        If Me.Criteria.Count > 0 Then
          DataSource.Criteria = Me.Criteria
        Else
          If DataSource.Criteria.Count > 0 Then
            ApplicationLogging.WriteLogEntry("Search.Criteria is not set, defaulting to DataSource.Criteria in Search::InitializeSearch", TraceEventType.Information, 4986)
          ElseIf DataSource.Clauses(0).Criteria.Count > 0 Then
            ApplicationLogging.WriteLogEntry("Search.Criteria is not set, defaulting to DataSource.Criteria in Search::InitializeSearch", TraceEventType.Information, 4987)
            DataSource.Criteria = DataSource.Clauses(0).Criteria
          Else
            ApplicationLogging.WriteLogEntry("No search criteria available to initialize search", TraceEventType.Warning, 5986)
          End If

        End If

        Dim lstrSQL As String
        If Args.UseDocumentValuesInCriteriaValues Then
          If Args.Document Is Nothing Then
            ApplicationLogging.WriteLogEntry(String.Format("Error in {0}::CSearch::InitializeSearch: 'The Document property of the SearchArgs is not set.'", Me.GetType.Name))
            Throw New InvalidOperationException("The Document property of the SearchArgs is not set.")
          End If
          lstrSQL = Me.DataSource.BuildSQLString(Args.Document, Args.VersionIndex, lstrErrorMessage)
        Else
          lstrSQL = Me.DataSource.BuildSQLString(lstrErrorMessage)
        End If


        If lstrErrorMessage.Length > 0 Then
          ApplicationLogging.WriteLogEntry(String.Format("Error in {0}::CSearch::InitializeSearch: '{1}'", Me.GetType.Name, lstrErrorMessage))
          Throw New ApplicationException("Error Creating SQL Statement: " & lstrErrorMessage)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CSearch::InitializeSearch(Args)", Me.GetType.Name))
        ' Rethrow the exception to the caller
        Throw
      End Try

    End Sub

    Protected Overridable Sub InitializeSearch()

      Dim lstrErrorMessage As String = ""

      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        ' Make sure the Provider property is set.
        If Provider Is Nothing Then
          Throw New InvalidOperationException("The Provider property is not set, cannot initialize the search")
        End If

        ' Initialize the DataSource with the correct QueryTarget and IDColumn

        If DataSource.QueryTarget.Length > 0 AndAlso DataSource.QueryTarget <> DOCUMENT_QUERY_TARGET Then
          'InitializeDataSource(Me.DataSource.ResultColumns(0), DataSource.QueryTarget)
          If (Me.DataSource.SourceColumn <> String.Empty) Then
            InitializeDataSource(Me.DataSource.SourceColumn, DataSource.QueryTarget)
          Else
            InitializeDataSource(Me.DataSource.ResultColumns(0), DataSource.QueryTarget)
          End If

        Else
          InitializeDataSource(ID_COLUMN, DOCUMENT_QUERY_TARGET)
        End If

        DataSource.Criteria = Me.Criteria

        DataSource.SQLStatement = Me.DataSource.BuildSQLString(lstrErrorMessage)

        If lstrErrorMessage.Length > 0 Then
          Throw New ApplicationException("Error Creating SQL Statement: " & lstrErrorMessage)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CSearch::InitializeSearch", Me.GetType.Name))
        ' Rethrow the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Sets the QueryTarget and SourceColumn of the member DataSource
    ''' </summary>
    ''' <param name="lpId_Column">Value to set for the SourceColumn</param>
    ''' <param name="lpQueryTarget">Value to set for the QueryTarget</param>
    ''' <remarks></remarks>
    Protected Overridable Sub InitializeDataSource(ByVal lpId_Column As String, ByVal lpQueryTarget As String)
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        With Me.DataSource
          .QueryTarget = lpQueryTarget
          .SourceColumn = lpId_Column
        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CSearch::InitializeDataSource('{1}', '{2}'", Me.GetType.Name, lpId_Column, lpQueryTarget))
      End Try
    End Sub

    Public Overridable Function GetSearchResultSetFromDataTable(ByVal lpDataTable As System.Data.DataTable) As Core.SearchResultSet
      Try

        Dim lobjSearchResult As Core.SearchResult
        Dim lobjSearchResultSet As New Core.SearchResultSet
        Dim lobjDataItem As Data.DataItem

        If lpDataTable Is Nothing Then
          Throw New ArgumentException("The data table is invalid")
        End If

        For Each lobjDataRow As System.Data.DataRow In lpDataTable.Rows
          lobjSearchResult = New SearchResult
          lobjSearchResult.ID = DataSource.SourceColumn


          For lintColumnCounter As Integer = 0 To lobjDataRow.ItemArray.Length - 1

            lobjDataItem = New Data.DataItem(lpDataTable.Columns(lintColumnCounter).ColumnName,
                                             lobjDataRow.ItemArray(lintColumnCounter))
            lobjSearchResult.Values.Add(lobjDataItem)

            If String.Compare(lobjDataItem.Name, DataSource.SourceColumn) = 0 Then
              lobjSearchResult.ID = lobjDataItem.Value
            End If
          Next

          lobjSearchResultSet.Results.Add(lobjSearchResult)

        Next

        Return lobjSearchResultSet

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try

    End Function

    Protected Sub RaiseSearchUpdateEvent(ByVal sender As Object, ByVal e As SearchEventArgs)

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      RaiseEvent SearchUpdate(sender, e)

    End Sub

    Protected Sub RaiseSearchCompleteEvent(ByVal sender As Object, ByVal e As SearchEventArgs)

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      RaiseEvent SearchComplete(sender, e)

    End Sub

#End Region

#Region "Private Methods"

    Private Sub InitializeDefaultResultColumns()
      Try
        mobjDefaultResultColumns = New ResultColumns

        If String.IsNullOrEmpty(DefaultDelimitedResultColumns) OrElse DefaultDelimitedResultColumns.Contains(",") = False Then
          Exit Sub
        End If

        Dim lstrSplitResultColumns As String() = DefaultDelimitedResultColumns.Split(",")
        Dim lstrCleanResultColumn As String = Nothing

        For Each lstrResultColumn As String In lstrSplitResultColumns
          lstrCleanResultColumn = lstrResultColumn.Trim(" ")
          If mobjDefaultResultColumns.Contains(lstrCleanResultColumn) = False Then
            mobjDefaultResultColumns.Add(lstrCleanResultColumn)
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: free other state (managed objects).
          mobjDataSource.Dispose()
        End If

        ' DISPOSETODO: free your own state (unmanaged objects).
        ' TODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

  End Class

End Namespace
