'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Search

  ''' <summary>
  ''' A search object that can be saved and executed at a later 
  ''' time which requires no user input parameters before execution
  ''' </summary>
  ''' <remarks>Serialized instances should use the 
  ''' file extension .ssf (Stored Search File)</remarks>
  Public Class StoredSearch
    Inherits Search
    Implements ISerialize

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const STORED_SEARCH_FILE_EXTENSION As String = "ssf"

#End Region

#Region "Class Variables"

    Private mstrOriginalFilePath As String = String.Empty
    Private mobjDocumentClasses As New ObservableCollection(Of DocumentClass)
    Private WithEvents mobjResultFormatters As FormatterItems = Nothing

#End Region

#Region "Public Properties"

    Public Overridable Property DocumentClasses() As ObservableCollection(Of DocumentClass)
      Get
        Try
          Return mobjDocumentClasses
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As ObservableCollection(Of DocumentClass))
        Try
          mobjDocumentClasses = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlAttribute()>
    Public Shadows Property DisplayName As String
      Get
        Try
          If MyBase.DisplayName Is Nothing OrElse MyBase.DisplayName.Length = 0 Then
            If OriginalFilePath.Length > 0 Then
              MyBase.DisplayName = IO.Path.GetFileNameWithoutExtension(OriginalFilePath)
            End If
          End If
          Return MyBase.DisplayName
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          MyBase.DisplayName = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property OriginalFilePath() As String
      Get
        Try
          Return mstrOriginalFilePath
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property ResultFormatters() As FormatterItems
      Get
        Try
          If mobjResultFormatters Is Nothing Then
            mobjResultFormatters = New FormatterItems
          End If
          Return mobjResultFormatters
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As FormatterItems)
        Try
          mobjResultFormatters = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overrides Property ResultColumns As Data.ResultColumns
      Get
        Try
          Return MyBase.ResultColumns
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Data.ResultColumns)
        Try
          MyBase.ResultColumns = value

          ' If this is not a result of opening an existing search we 
          ' need to syncronize the result columns to the result formatters.
          If Helper.CallStackContainsMethodName("AssignObjectProperties") = False Then
            For Each lstrResultColumn As String In value
              If ResultFormatters.Contains(lstrResultColumn) = False Then
                ResultFormatters.Add(lstrResultColumn)
              End If
            Next
          End If
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Constructs a StoredSearch object by deserializing the XML in the specified
    ''' path.
    ''' </summary>
    ''' <param name="lpXMLFilePath">A fully qualified XML file path for the serialized StoredSearch file.</param>
    Public Sub New(ByVal lpXMLFilePath As String)
      Dim lobjXMLDocument As New Xml.XmlDocument
      Try
        mstrOriginalFilePath = lpXMLFilePath
        lobjXMLDocument.Load(lpXMLFilePath)
        LoadFromXmlDocument(lobjXMLDocument)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("StoredSearch::New(lpXMLFilePath:'{0}')", lpXMLFilePath))
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      Try
        LoadFromXmlDocument(lpXML)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    'Public Sub AddCriterionFilter(ByVal lpCriterion As Criterion, ByVal lpClause As Clause)
    '  Try
    '    lpClause.Criteria.Add(lpCriterion)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Sub AddDocumentClassFilter(ByVal lpDocumentClass As DocumentClass)
      Try
        DocumentClasses.Add(lpDocumentClass)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function Execute() As Core.SearchResultSet

      Try

        Dim lobjSearchInterface As ISearch = Nothing

        Dim lobjSearchArgs As New Arguments.SearchArgs

        With lobjSearchArgs

        End With

        ' NOTE: We currently are only supporting a single content source, 
        '       federated searches will be properly implemented at a future date.
        Dim lobjContentSource As ContentSource = ContentSources.FirstOrDefault()
        If lobjContentSource Is Nothing Then
          Throw New Exceptions.InvalidContentSourceException("Unable to execute search, no content source is available.")
        End If

        'For Each lobjContentSource As ContentSource In ContentSources()

        If lobjContentSource.State = ProviderConnectionState.Disconnected Then
          lobjContentSource.Provider.Connect(lobjContentSource)
        End If

        lobjSearchInterface = lobjContentSource.Provider.CreateSearch()

        ' <Added by: Ernie at: 3/27/2012-11:07:15 AM on machine: ERNIE-M4400>
        If Me.PassThroughSqlStatements.Count > 0 Then
          If lobjContentSource.Provider.SupportsInterface(ProviderClass.SQLPassThroughSearch) Then
            Dim lobjPassThroughSearcher As ISQLPassThroughSearch = CType(lobjSearchInterface, ISQLPassThroughSearch)

            ' NOTE: We are currently only supporting a single SQL statement.
            '       Federated or union passthrough may be properly supported at a future date.
            Me.mobjSearchResultSet = lobjPassThroughSearcher.Execute(Me.PassThroughSqlStatements(0))
            Return Me.mobjSearchResultSet
          Else
            Throw New Exceptions.ProviderInterfaceNotImplementedException(lobjContentSource.Provider, ProviderClass.SQLPassThroughSearch)
          End If
        Else
          ' </Added by: Ernie at: 3/27/2012-11:07:15 AM on machine: ERNIE-M4400>
          With lobjSearchInterface
            If Clauses.Count = 1 Then
              .Criteria = Clauses.Item(0).Criteria
            End If
            .DataSource.Clauses = Clauses
            .DataSource.LimitResults = MaxResults
            .DataSource.ResultColumns = ResultColumns
            .DataSource.OrderBy = OrderItems
            .DataSource.QueryTarget = QueryTarget
            If (IdColumn <> String.Empty) Then
              .DataSource.SourceColumn = IdColumn
            End If
          End With

          Me.mobjSearchResultSet = lobjSearchInterface.Execute

          Return Me.mobjSearchResultSet
        End If

        'Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Me.mobjSearchResultSet = New SearchResultSet 'Same as initialized in base class
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Private Methods"

    Private Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
      Try
        ' In case we have a value for OriginalFilePath we may need to cache it here.
        Dim lstrOrginalFilePath As String = OriginalFilePath
        Dim lobjStoredSearch As StoredSearch = Deserialize(lpXML)
        Helper.AssignObjectProperties(lobjStoredSearch, Me)

        ' If we have a cached value for original file path that was wiped out by the serialization, replace it here.
        If String.Compare(lstrOrginalFilePath, OriginalFilePath, True) <> 0 Then
          mstrOriginalFilePath = lstrOrginalFilePath
        End If
        If DisplayName.Length = 0 AndAlso OriginalFilePath.Length > 0 Then
          DisplayName = IO.Path.GetFileNameWithoutExtension(OriginalFilePath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    Private Sub mobjResultFormatters_CollectionChanged(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles mobjResultFormatters.CollectionChanged
      Try
        ' Syncronize any new resultformatters to the collection of order items
        Select Case e.Action
          Case Specialized.NotifyCollectionChangedAction.Add
            For Each lobjFormatterItem As FormatterItem In e.NewItems
              If ResultColumns.Contains(lobjFormatterItem.FormatterItem.Name) = False Then
                ResultColumns.Add(lobjFormatterItem.FormatterItem.Name)
              End If
            Next
          Case Specialized.NotifyCollectionChangedAction.Remove
            For Each lobjFormatterItem As FormatterItem In e.OldItems
              If ResultColumns.Contains(lobjFormatterItem.FormatterItem.Name) = False Then
                ResultColumns.Remove(lobjFormatterItem.FormatterItem.Name)
              End If
            Next
        End Select
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

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
        Return STORED_SEARCH_FILE_EXTENSION
      End Get
    End Property

    ''' <summary>
    ''' Instantiate from an XML file.
    ''' </summary>
    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        mstrOriginalFilePath = lpFilePath
        Dim lblnReturnValue As Boolean = Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

        Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

        If DisplayName.Length = 0 AndAlso OriginalFilePath.Length > 0 Then
          DisplayName = IO.Path.GetFileNameWithoutExtension(OriginalFilePath)
        End If

        Return lblnReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try

        If lpFileExtension.Length = 0 Then
          ' No override was provided
          If lpFilePath.EndsWith(DefaultFileExtension) = False Then
            lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DefaultFileExtension
          End If

        End If

        ' If no display name was set previously, set it to the requested file name
        If DisplayName.Length = 0 Then
          DisplayName = IO.Path.GetFileNameWithoutExtension(lpFilePath)
        End If

        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace