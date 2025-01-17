'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Data
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Namespace Search

  ''' <summary>
  ''' Base search class from which StoredSearch and SearchTemplate are derived
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class Search
    Implements IDisposable
    Implements IDisplayable

#Region "Class Variables"

    Private mstrDisplayName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrApplicationName As String = String.Empty
    Private mobjConnectionStrings As New List(Of String)
    Private mobjContentSources As ContentSources
    Private mobjPassThroughSqlStatements As New List(Of String)
    Private mobjClauses As New Clauses
    Private mobjResultColumns As New ResultColumns
    Private mobjOrderItems As New OrderItems
    Private mobjFolderFilters As New FolderFilters
    Friend mobjSearchResultSet As New Core.SearchResultSet
    Private mlngMaxResults As Long
    Private mstrQueryTarget As String = String.Empty
    Public Const SOURCE_CONNECTIONSTRING_FIELD_NAME As String = "SourceConnString"
    Private mstrIdColumn As String = String.Empty
    Private mstrFirstContentSourceName As String = String.Empty
    Private mblnAllowOfflineEdits As Boolean = False

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the display name to be used for the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property DisplayName() As String Implements IDescription.Name, INamedItem.Name, IDisplayable.DisplayName
      Get
        Try
          Return mstrDisplayName
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrDisplayName = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a description for the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Description() As String Implements IDescription.Description
      Get
        Try
          Return mstrDescription
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrDescription = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the application name associated with the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ApplicationName() As String
      Get
        Try
          Return mstrApplicationName
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrApplicationName = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the connection string for the 
    ''' content source that will execute the search
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionStrings() As List(Of String)
      Get
        Try
          Return mobjConnectionStrings
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As List(Of String))
        Try
          mobjConnectionStrings = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    '''     Determines whether or not offline edits will be allowed for this search.
    ''' </summary>
    ''' <remarks>
    '''     In order to enable offline edits, all of the content source and document class details will need to be saved with the search.  The file will be MUCH larger.
    ''' </remarks>
    Public Property AllowOfflineEdits As Boolean
      Get
        Try
          Return mblnAllowOfflineEdits
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          mblnAllowOfflineEdits = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property FirstConnectionString As String
      Get
        Try
          If ConnectionStrings Is Nothing Then
            Return String.Empty
          End If
          If ConnectionStrings.Count > 0 Then
            Return ConnectionStrings(0)
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property FirstContentSourceName As String
      Get
        Try
          If String.IsNullOrEmpty(mstrFirstContentSourceName) Then
            mstrFirstContentSourceName = ContentSource.GetNameFromConnectionString(FirstConnectionString)
          End If
          Return mstrFirstContentSourceName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public ReadOnly Property ContentSources() As ContentSources
      Get
        Try

          ' If the collection is empty then we need to initialize it
          If mobjContentSources Is Nothing OrElse mobjContentSources.Count = 0 Then
            InitializeContentSources()
          End If

          Return mobjContentSources

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property PassThroughSqlStatements As List(Of String)
      Get
        Try
          Return mobjPassThroughSqlStatements
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As List(Of String))
        Try
          mobjPassThroughSqlStatements = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a collection of search clauses
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Each clause may contain one 
    ''' or more search criteria</remarks>
    Public Property Clauses() As Clauses
      Get
        Try
          Return mobjClauses
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Clauses)
        Try
          mobjClauses = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable Property ResultColumns() As ResultColumns
      Get
        Try
          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          Return mobjResultColumns
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As ResultColumns)
        Try
          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          mobjResultColumns = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of order items
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OrderItems() As OrderItems
      Get
        Try
          Return mobjOrderItems
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As OrderItems)
        Try
          mobjOrderItems = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a collection of folder filters for 
    ''' limiting the scope of a search to within one or more folders
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FolderFilters() As FolderFilters
      Get
        Try
          Return mobjFolderFilters
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As FolderFilters)
        Try
          mobjFolderFilters = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the limit on the number of search results
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property MaxResults() As Long
      Get
        Try
          Return mlngMaxResults
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Long)
        Try
          mlngMaxResults = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the set of search results from the last search execution
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SearchResultSet() As Core.SearchResultSet
      Get
        Try
          Return mobjSearchResultSet
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets/Sets the target of the query (typically document class or table name)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property QueryTarget() As String
      Get
        Try
          Return mstrQueryTarget
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrQueryTarget = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property IdColumn() As String
      Get
        Try
          Return mstrIdColumn
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrIdColumn = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Private Properties"

    Protected ReadOnly Property IsDisposed() As Boolean
      Get
        Try
          Return disposedValue
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Executes the current search
    ''' </summary>
    ''' <returns>A SearchResults collection</returns>
    ''' <remarks></remarks>
    Public MustOverride Function Execute() As Core.SearchResultSet

    Public Sub AddContentSource(ByVal lpContentSource As ContentSource)
      Try

        If lpContentSource Is Nothing Then
          Throw New ArgumentNullException
        End If

        If mobjContentSources.Contains(lpContentSource) = False Then
          mobjContentSources.Add(lpContentSource)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If Not String.IsNullOrEmpty(DisplayName) Then
          Return String.Format("{0} ({1})", DisplayName, FirstContentSourceName)
        Else
          Return String.Format("Unamed search against '{0}'", DisplayName, FirstContentSourceName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub InitializeContentSources()
      Try

        Dim lobjContentSource As ContentSource = Nothing

        If mobjContentSources Is Nothing Then
          ' Initialize the collection
          mobjContentSources = New ContentSources
        Else
          ' Make sure the collection is empty
          mobjContentSources.Clear()
        End If

        ' Loop through each connection string and make sure we have the content source initialized
        For Each lstrConnectionString As String In ConnectionStrings

          ' Try to create a content source using the connection string
          lobjContentSource = New ContentSource(lstrConnectionString)

          ' Add the content source to the collection
          mobjContentSources.Add(lobjContentSource)

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDisposable Implementation"

    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: free other state (managed objects).
          End If

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

#End Region

  End Class

End Namespace

