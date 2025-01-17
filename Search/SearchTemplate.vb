'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities


#End Region

Namespace Search

  ''' <summary>
  ''' A search object that can be saved and executed at a later time but requires user input parameters before execution
  ''' </summary>
  ''' <remarks></remarks>
  Public Class SearchTemplate
    Inherits StoredSearch
    Implements ISerialize

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const SEARCH_TEMPLATE_FILE_EXTENSION As String = "stf"

#End Region

#Region "Public Enumerations"

#End Region

#Region "Class Variables"

    ''' <summary>
    ''' Stores the curently selected document 
    ''' classes that will be used to filter the search.
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents mobjDocumentClassFilters As New FilterItems(Of RepositoryObjectClass)(New RepositoryObjectClasses)

    ''' <summary>
    ''' Stores the currently selected properties that will be used to filter the search.
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents mobjPropertyFilters As New FilterItems(Of ClassificationProperty)(New ClassificationProperties)

#End Region

#Region "Public Properties"

    Public Property DocumentClassFilters() As FilterItems(Of RepositoryObjectClass)
      Get
        Return mobjDocumentClassFilters
      End Get
      Set(ByVal value As FilterItems(Of RepositoryObjectClass))
        mobjDocumentClassFilters = value
      End Set
    End Property

    Public Property PropertyFilters() As FilterItems(Of ClassificationProperty)
      Get
        Return mobjPropertyFilters
      End Get
      Set(ByVal value As FilterItems(Of ClassificationProperty))
        mobjPropertyFilters = value
      End Set
    End Property

#End Region

#Region "Constructors"

#End Region

#Region "Public Methods"

    Public Overloads Sub AddDocumentClassFilter(ByVal lpDocumentClass As TemplatedDocumentClass)
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

        ' TODO: Implement this method
        Return New SearchResultSet

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

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
    Public Overloads ReadOnly Property DefaultFileExtension() As String
      Get
        Return SEARCH_TEMPLATE_FILE_EXTENSION
      End Get
    End Property

#End Region

  End Class

End Namespace

