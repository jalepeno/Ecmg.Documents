' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ItemPermission.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 3:34:00 PM
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

  <DebuggerDisplay("{DebuggerIdentifier(),nq}"), XmlRoot("ItemPermission")>
  Public Class ItemPermission
    Implements IPermission

#Region "Class Variables"

    Private mstrPrincipalName As String = String.Empty
    Private mobjAccess As IAccessMask = Nothing
    Private menuAccessType As AccessType = Security.AccessType.Allow
    Private menuSource As PermissionSource = Security.PermissionSource.Direct
    Private menuPrincipalType As PrincipalType = Security.PrincipalType.User

#End Region

#Region "IPermission Implementation"

    Public Property PrincipalName As String Implements IPermission.PrincipalName
      Get
        Return mstrPrincipalName
      End Get
      Set(value As String)
        mstrPrincipalName = value
      End Set
    End Property

    Public Property Access As AccessRight
      Get
        Return mobjAccess
      End Get
      Set(value As AccessRight)
        mobjAccess = value
      End Set
    End Property

    <XmlIgnore()>
    Public Property IAccess As IAccessMask Implements IPermission.Access
      Get
        Return mobjAccess
      End Get
      Set(value As IAccessMask)
        mobjAccess = value
      End Set
    End Property

    Public Property AccessType As AccessType Implements IPermission.AccessType
      Get
        Return menuAccessType
      End Get
      Set(value As AccessType)
        menuAccessType = value
      End Set
    End Property

    Public Property Source As PermissionSource Implements IPermission.PermissionSource
      Get
        Return menuSource
      End Get
      Set(value As PermissionSource)
        menuSource = value
      End Set
    End Property

    Public Property PrincipalType As PrincipalType Implements IPermission.PrincipalType
      Get
        Return menuPrincipalType
      End Get
      Set(value As PrincipalType)
        menuPrincipalType = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpPrincipalName As String, lpAccessType As AccessType, lpAccess As IAccessMask, lpSource As PermissionSource)
      Try
        PrincipalName = lpPrincipalName
        AccessType = lpAccessType
        Access = lpAccess
        Source = lpSource
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        Return Me.GetType.Name
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        If Me.Access IsNot Nothing Then
          lobjIdentifierBuilder.AppendFormat("{0}: {1} - {2} ~ {3}",
                                             Me.PrincipalType,
                                             Me.PrincipalName,
                                             Me.Source.ToString(),
                                             Me.Access.DebuggerIdentifier())
        Else
          lobjIdentifierBuilder.AppendFormat("{0}: {1} - {2} ~ (Access Not Defined)",
                                             Me.PrincipalType,
                                             Me.PrincipalName,
                                             Me.Source.ToString())
        End If

        Return lobjIdentifierBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function
#End Region

  End Class

End Namespace