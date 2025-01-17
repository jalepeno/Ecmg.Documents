'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Base argument class for Cts operations that work with a document
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class DocumentEventArgs
    Inherits CtsEventArgs

#Region "Class Variables"

    Private mobjDocument As Document
    Private mstrDocumentId As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The document for which the event occured
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Document() As Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Document)
        mobjDocument = value
      End Set
    End Property

    Public Property DocumentId As String
      Get
        Return mstrDocumentId
      End Get
      Set(ByVal value As String)
        mstrDocumentId = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document)
      Me.New(lpDocument, "", Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpEventDescription As String)
      Me.New(lpDocument, lpEventDescription, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpTime As DateTime)
      Me.New(lpDocument, "", lpTime)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpEventDescription As String, ByVal lpTime As DateTime)
      Document = lpDocument
      EventDescription = lpEventDescription
      Time = lpTime
    End Sub

    Public Sub New(ByVal lpDocumentId As String)
      Me.New(lpDocumentId, "", Now)
    End Sub

    Public Sub New(ByVal lpDocumentId As String, ByVal lpEventDescription As String)
      Me.New(lpDocumentId, lpEventDescription, Now)
    End Sub

    Public Sub New(ByVal lpDocumentId As String, ByVal lpTime As DateTime)
      Me.New(lpDocumentId, "", lpTime)
    End Sub

    Public Sub New(ByVal lpDocumentId As String, ByVal lpEventDescription As String, ByVal lpTime As DateTime)
      DocumentId = lpDocumentId
      EventDescription = lpEventDescription
      Time = lpTime
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        lobjReturnBuilder.AppendFormat("Doc={0}; Time={1}; Desc={2}; Success={3}; Message={4}",
                                       Document.DebuggerIdentifier, Time, EventDescription,
                                       OperationSucceeded.ToString, Message)

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace