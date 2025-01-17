' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  AccessRight.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:53:30 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Security

  <XmlInclude(GetType(AccessLevel))>
  Public Class AccessRight
    Inherits EnumeratedValue
    Implements IAccessRight

#Region "Class Constants"

    Private Const PARENT_NAME As String = "AccessRight"

#End Region

#Region "Class Variables"

    Private mobjPermissionList As IPermissionList = New PermissionList

#End Region

#Region "IAccessMask Implementation"

    <XmlIgnore()>
    Public Property Permissions As IPermissionList Implements IAccessRight.PermissionList
      Get
        Return mobjPermissionList
      End Get
      Set(value As IPermissionList)
        mobjPermissionList = value
      End Set
    End Property

    Public Property PermissionList As PermissionList
      Get
        Return mobjPermissionList
      End Get
      Set(value As PermissionList)
        mobjPermissionList = value
      End Set
    End Property

    Public Overridable ReadOnly Property MaskType As PermissionType Implements IAccessMask.Type
      Get
        Return PermissionType.Right
      End Get
    End Property

    Public Shadows Property Name As String Implements IAccessMask.Name
      Get
        Return MyBase.Name
      End Get
      Set(value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Shadows Property Value As Nullable(Of Integer) Implements IAccessMask.Value
      Get
        Return MyBase.Value
      End Get
      Set(value As Nullable(Of Integer))
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.new()
      Try
        ParentName = PARENT_NAME
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpName As String, lpValue As Integer)
      MyBase.New(PARENT_NAME, lpName, lpValue)
    End Sub

#End Region

  End Class

End Namespace