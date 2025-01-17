'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters necessary for the CSearch.Execute method.
  ''' </summary>
  ''' <remarks>Used as the sole parameter of the CSearch.Execute method.</remarks>
  Public Class SearchArgs

#Region "Class Variables"

    Private mblnUseDocumentValuesInCriteriaValues As Boolean
    Private mobjDocument As Document
    Private mintVersionIndex As Integer
    Private mstrErrorMessage As String

#End Region

#Region "Public Properties"

    Public ReadOnly Property Document() As Document
      Get
        Return mobjDocument
      End Get
    End Property

    Public ReadOnly Property UseDocumentValuesInCriteriaValues() As Boolean
      Get
        Return mblnUseDocumentValuesInCriteriaValues
      End Get
    End Property

    Public ReadOnly Property VersionIndex() As Integer
      Get
        Return mintVersionIndex
      End Get
    End Property

    Public Property ErrorMessage() As String
      Get
        Return mstrErrorMessage
      End Get
      Set(ByVal value As String)
        mstrErrorMessage = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Me.New(Nothing, False, 0)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpUseDocumentValuesInCriteriaValues As Boolean)
      Me.New(lpDocument, lpUseDocumentValuesInCriteriaValues, 0)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpUseDocumentValuesInCriteriaValues As Boolean, ByVal lpVersionIndex As Integer)
      mobjDocument = lpDocument
      mblnUseDocumentValuesInCriteriaValues = lpUseDocumentValuesInCriteriaValues
      mintVersionIndex = lpVersionIndex
    End Sub

#End Region

  End Class
End Namespace