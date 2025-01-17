'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentUnFiled Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentUnFiledEventArgs
    Inherits DocumentFilingEventArgs

#Region "Constructors"

    Public Sub New(ByVal lpId As String, ByVal lpFolderPath As String)
      Me.New(lpId, lpFolderPath, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpFolderPath As String)
      Me.New(lpDocument, lpFolderPath, Now)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpFolderPath As String, ByVal lpTime As DateTime)
      MyBase.New(lpId, lpFolderPath, lpTime)
      Me.EventDescription = "DocumentUnFiled"
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpFolderPath As String, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpFolderPath, lpTime)
      Me.EventDescription = "DocumentUnFiled"
    End Sub

#End Region

  End Class
End Namespace