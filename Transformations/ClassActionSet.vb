'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>
  ''' Represents a set of transformation actions 
  ''' to perform for a specific document class.
  ''' </summary>
  ''' <remarks></remarks>
  <XmlRoot("DocumentClassActions")>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ClassActionSet

#Region "Class Variables"

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The document class for which the actions will apply
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute("DocumentClass")>
    Public Property DocumentClass As String

    ''' <summary>
    ''' The set of actions to perform for the current document class
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Actions As New Actions

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpDocumentClassName As String)
      Try
        DocumentClass = lpDocumentClassName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpDocumentClassName As String, lpActions As Actions)
      Try
        DocumentClass = lpDocumentClassName
        Actions = lpActions
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Friend Function DebuggerIdentifier() As String
      Try

        If String.IsNullOrEmpty(DocumentClass) Then
          If Actions.Count <> 1 Then
            Return String.Format("Document Class Not Set: {0} Actions", Actions.Count)
          Else
            Return String.Format("Document Class Not Set: {0} Action", Actions.Count)
          End If

        Else
          If Actions.Count <> 1 Then
            Return String.Format("{0}: {1} Actions", DocumentClass, Actions.Count)
          Else
            Return String.Format("{0}: {1} Action", DocumentClass, Actions.Count)
          End If
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