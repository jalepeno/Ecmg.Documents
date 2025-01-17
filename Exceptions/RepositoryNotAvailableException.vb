'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Exceptions
  ''' <summary>
  ''' Thrown when the repository is not currently available
  ''' </summary>
  ''' <remarks></remarks>
  Public Class RepositoryNotAvailableException
    Inherits CtsException

#Region "Class Variables"

    Private mstrRepositoryName As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property RepositoryName() As String
      Get
        Return mstrRepositoryName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpRepositoryName As String)
      MyBase.New(InitializeStandardMessage(lpRepositoryName))
      mstrRepositoryName = lpRepositoryName
    End Sub

    Public Sub New(ByVal lpRepositoryName As String,
                   ByVal message As String)
      MyBase.New(message)
      mstrRepositoryName = lpRepositoryName
    End Sub

    Public Sub New(ByVal lpRepositoryName As String,
                   ByVal message As String,
                   ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrRepositoryName = lpRepositoryName
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function InitializeStandardMessage(ByVal lpRepositoryName As String) As String
      Try
        Return String.Format("The repository '{0}' is currently unavailable", lpRepositoryName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace