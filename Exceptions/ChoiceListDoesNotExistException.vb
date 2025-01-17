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
  Public Class ChoiceListDoesNotExistException
    Inherits ChoiceListException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ChoiceListDoesNotExistException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceListName As String)
      Me.New(choiceListName,
             FormatMessage(choiceListName),
             Nothing)
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListDoesNotExistException with the specified choice list name.
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
    ''' Creates a new ChoiceListDoesNotExistException with the specified choice list name.
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

    Private Shared Function FormatMessage(ByVal choiceListName As String) As String
      Try
        Return String.Format("No choice list by the name '{0}' could be found", choiceListName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace