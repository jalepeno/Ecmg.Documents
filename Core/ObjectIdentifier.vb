'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>
  ''' Flexible class for object identification.
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ObjectIdentifier

#Region "Public Enumerations"

    ''' <summary>Specifies the way an identifier specifies an object</summary>
    Public Enum IdTypeEnum
      ''' <summary>Specifies an object via a unique identification value</summary>
      ID = 0
      ''' <summary>Specifies an object via its name</summary>
      Name = 1
    End Enum

    Public Enum ObjectTypeEnum
      Document = 0
      Folder = 1
      Custom = 2
      Other = 3
    End Enum

#End Region

#Region "Class Variables"

    Private menuIdType As IdTypeEnum
    Private menuObjectType As ObjectTypeEnum
    Private mstrIdentifierValue As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property IdentifierValue() As String
      Get
        Return mstrIdentifierValue
      End Get
    End Property

    Public ReadOnly Property IdentifierType() As IdTypeEnum
      Get
        Return menuIdType
      End Get
    End Property

    Public ReadOnly Property ObjectType As ObjectTypeEnum
      Get
        Return menuObjectType
      End Get
    End Property

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If IdentifierValue.Length > 0 Then
          Return String.Format("{0} = {1}", IdentifierType, IdentifierValue)
        Else
          Return String.Format("{0} = <value not set>", IdentifierType)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Finalizer"

    Protected Overrides Sub Finalize()
      MyBase.Finalize()
    End Sub

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpIdentifierValue As String, ByVal lpIdentifierType As IdTypeEnum)
      Try
        menuIdType = lpIdentifierType
        mstrIdentifierValue = lpIdentifierValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpIdentifierValue As String, ByVal lpObjectType As ObjectTypeEnum)
      Try
        menuIdType = IdTypeEnum.ID
        mstrIdentifierValue = lpIdentifierValue
        menuObjectType = lpObjectType
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpIdentifierValue As String, ByVal lpIdentifierType As IdTypeEnum, ByVal lpObjectType As ObjectTypeEnum)
      Try
        menuIdType = lpIdentifierType
        mstrIdentifierValue = lpIdentifierValue
        menuObjectType = lpObjectType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace