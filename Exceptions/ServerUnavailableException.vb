'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ServerUnavailableException.vb
'   Description :  [type_description_here]
'   Created     :  12/18/2013 11:11:09 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  Public Class ServerUnavailableException
    Inherits InvalidOperationException

#Region "Class Variables"

    Private mstrServerName As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ServerName() As String
      Get
        Return mstrServerName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpServerName As String)
      MyBase.New(InitializeStandardMessage(lpServerName))
      mstrServerName = lpServerName
    End Sub

    Public Sub New(ByVal lpServerName As String,
                   ByVal message As String)
      MyBase.New(message)
      mstrServerName = lpServerName
    End Sub

    Public Sub New(ByVal lpServerName As String,
                   ByVal message As String,
                   ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrServerName = lpServerName
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function InitializeStandardMessage(ByVal lpServerName As String) As String
      Try
        Return String.Format("The server '{0}' is currently unavailable", lpServerName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace