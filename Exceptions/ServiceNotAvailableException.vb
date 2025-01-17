' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ServiceNotAvailableException.vb
'  Description :  [type_description_here]
'  Created     :  8/17/2016 2:37:50 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions
  Public Class ServiceNotAvailableException
    Inherits CtsException

#Region "Class Variables"

    Private mstrServiceName As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ServiceName() As String
      Get
        Return mstrServiceName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpServiceName As String)
      MyBase.New(InitializeStandardMessage(lpServiceName))
      mstrServiceName = lpServiceName
    End Sub

    Public Sub New(ByVal lpServiceName As String,
                   ByVal message As String)
      MyBase.New(message)
      mstrServiceName = lpServiceName
    End Sub

    Public Sub New(ByVal lpServiceName As String,
                   ByVal message As String,
                   ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrServiceName = lpServiceName
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function InitializeStandardMessage(ByVal lpServiceName As String) As String
      Try
        Return String.Format("The service '{0}' is currently unavailable", lpServiceName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace