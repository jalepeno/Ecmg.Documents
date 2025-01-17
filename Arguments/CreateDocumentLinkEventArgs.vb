' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Process.vb
'  Description :  [type_description_here]
'  Created     :  11/15/2012 3:48:31 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters necessary for the CreateDocumentLink method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the CreateDocumentLink method.</remarks>
  Public Class CreateDocumentLinkEventArgs
    Inherits DocumentLinkEventArgs

#Region "Constructors"

    Public Sub New(ByVal lpName As String,
               ByVal lpTargetUrl As String,
               ByVal lpMimeType As String)
      MyBase.New(lpName, lpTargetUrl, lpMimeType, lpMimeType, Nothing)
    End Sub

    Public Sub New(ByVal lpName As String,
            ByVal lpTargetUrl As String,
            ByVal lpMimeType As String,
            ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpName, lpTargetUrl, lpMimeType, lpMimeType, lpWorker)
    End Sub

    Public Sub New(ByVal lpName As String,
                   ByVal lpTargetUrl As String,
                   ByVal lpFileType As String,
                   ByVal lpMimeType As String,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpName, lpTargetUrl, lpFileType, lpMimeType, lpWorker)
    End Sub

#End Region

  End Class

End Namespace