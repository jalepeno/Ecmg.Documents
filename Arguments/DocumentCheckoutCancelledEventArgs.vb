'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentCheckOutCancelled Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentCheckoutCancelledEventArgs
    Inherits DocumentCheckedOutEventArgs

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String)
      Me.New(lpDocument, "", lpCheckedOutUserName, lpVersionId, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String)
      Me.New(lpDocument, lpDestinationPath, lpCheckedOutUserName, lpVersionId, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpDestinationPath, New String() {}, lpCheckedOutUserName, lpVersionId, lpTime)
      EventDescription = "DocumentCheckoutCancelled"
    End Sub

#End Region

  End Class
End Namespace