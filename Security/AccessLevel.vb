' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  AccessLevel.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:58:05 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Security

  Public Class AccessLevel
    Inherits AccessRight
    Implements IAccessLevel

#Region "Class Constants"

    Private Const PARENT_NAME As String = "AccessLevel"

#End Region

#Region "Class Variables"

    Private menuLevel As PermissionLevel
    Private mobjRights As New AccessRights

#End Region

#Region "Public Properties"

#Region "IAccessMask Implementation"

    Public Overrides ReadOnly Property MaskType As PermissionType
      Get
        Return PermissionType.Level
      End Get
    End Property

    Public Property Level As PermissionLevel Implements IAccessLevel.Level
      Get
        Return menuLevel
      End Get
      Set(value As PermissionLevel)
        menuLevel = value
      End Set
    End Property

    Public Shadows Property Name As String
      Get
        Return MyBase.Name
      End Get
      Set(value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Shadows Property Value As Nullable(Of Integer)
      Get
        Return MyBase.Value
      End Get
      Set(value As Nullable(Of Integer))
        MyBase.Value = value
      End Set
    End Property

#End Region

    Public Property Rights As AccessRights
      Get
        Return mobjRights
      End Get
      Set(value As AccessRights)
        mobjRights = value
      End Set
    End Property

    <XmlIgnore()>
    Public Property IRights As IAccessRights Implements IAccessLevel.Rights
      Get
        Return mobjRights
      End Get
      Set(value As IAccessRights)
        mobjRights = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(lpLevel As PermissionLevel)
      MyBase.New()
      Try
        Level = lpLevel
        Name = lpLevel.ToString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpName As String, lpValue As Integer)
      MyBase.New(lpName, lpValue)
      Try
        ParentName = PARENT_NAME
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpName As String, lpValue As Integer, lpRights As AccessRights)
      MyBase.New(lpName, lpValue)
      Try
        ParentName = PARENT_NAME
        Rights = lpRights
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        Select Case Rights.Count
          Case 0
            lobjIdentifierBuilder.AppendFormat("{0} - no rights defined", MyBase.DebuggerIdentifier)
          Case 1
            lobjIdentifierBuilder.AppendFormat("{0} - 1 right defined", MyBase.DebuggerIdentifier)
          Case Else
            lobjIdentifierBuilder.AppendFormat("{0} - {1} rights defined", MyBase.DebuggerIdentifier, Rights.Count)
        End Select

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace