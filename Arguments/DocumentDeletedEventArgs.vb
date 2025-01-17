'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentDeleted Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentDeletedEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mstrId As String

#End Region

#Region "Public Properties"

    Public ReadOnly Property Id() As String
      Get
        Return mstrId
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      Me.New(lpId, Now)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpTime As DateTime)
      MyBase.New(String.Empty, "DocumentDeleted", lpTime)
      mstrId = lpId
    End Sub

#End Region

  End Class
End Namespace