'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Exceptions

#End Region

''' <summary>
''' Contains the defined exceptions of the ChoiceLists namespace of the Cts framework
''' </summary>
Namespace Core.ChoiceLists.Exceptions

  ''' <summary>
  ''' Exception thrown by a ChoiceList or from methods associated with ChoiceList
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceListException
    Inherits CtsException

#Region "Class Variables"

    Private m_choiceListName As String

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The Document Identifier
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ChoiceListName() As String
      Get
        Return m_choiceListName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal choiceListName As String,
                   ByVal message As String)
      MyBase.New(message)
      m_choiceListName = choiceListName
    End Sub

    ''' <summary>
    ''' Creates a new ChoiceListException with the specified choice list name.
    ''' </summary>
    ''' <param name="choiceListName">The name of the choice list</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal choiceListName As String,
                   ByVal message As String,
                   ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      m_choiceListName = choiceListName
    End Sub

#End Region

  End Class
End Namespace