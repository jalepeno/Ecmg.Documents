'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

''' <summary>
''' Contains the defined exceptions of the ChoiceLists namespace of the Cts framework
''' </summary>
Namespace Core.ChoiceLists.Exceptions

  ''' <summary>
  ''' Exception thrown by a document or from methods associated with documents
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceListAlreadyExistsException
    Inherits ChoiceListException

#Region "Class Variables"

    Private mobjChoiceList As ChoiceList = Nothing
    Private mstrRepositoryName As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ChoiceList() As ChoiceList
      Get
        Return mobjChoiceList
      End Get
    End Property

    Public ReadOnly Property RepositoryName() As String
      Get
        Return mstrRepositoryName
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list.
    ''' </summary>
    ''' <param name="choiceList">The choice list found in the destination</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceList As ChoiceList)
      Me.New(choiceList.Name,
             FormatMessage(choiceList.Name),
             Nothing)
      mobjChoiceList = choiceList
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceListName As String, ByVal repositoryName As String)
      Me.New(choiceListName,
             FormatMessage(choiceListName, repositoryName),
             Nothing)
      mstrRepositoryName = repositoryName
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list.
    ''' </summary>
    ''' <param name="choiceList">The choice list found in the destination</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceList As ChoiceList, ByVal repositoryName As String)
      Me.New(choiceList.Name,
             FormatMessage(choiceList.Name, repositoryName),
             Nothing)
      mobjChoiceList = choiceList
      mstrRepositoryName = repositoryName
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceListName As String)
      Me.New(choiceListName,
             FormatMessage(choiceListName),
             Nothing)
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <param name="innerException"></param>
    ''' <remarks></remarks>
    Protected Sub New(ByVal choiceListName As String,
                      ByVal innerException As Exception)
      Me.New(choiceListName,
             FormatMessage(choiceListName),
             innerException)
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListAlreadyExistsException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceListName As String,
                   ByVal message As String,
                   ByVal innerException As System.Exception)
      MyBase.New(choiceListName, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function FormatMessage(ByVal choiceListName As String,
                                          Optional ByVal repositoryName As String = "") As String
      Try
        If repositoryName.Length > 0 Then
          Return String.Format("A choice list by the name '{0}' already exists in the repository '{1}'", choiceListName, repositoryName)
        Else
          Return String.Format("A choice list by the name '{0}' already exists in the repository", choiceListName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace