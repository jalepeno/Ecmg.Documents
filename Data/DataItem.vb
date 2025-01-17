'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Data

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class DataItem
    Inherits Core.Value

#Region "Class Variables"

    Private menuType As Core.PropertyType = Core.PropertyType.ecmString
    Private mstrName As String = String.Empty
    Private mblnIsKey As Boolean

#End Region

#Region "Public Properties"

    Public Property Type() As Core.PropertyType
      Get
        Return menuType
      End Get
      Set(ByVal value As Core.PropertyType)
        menuType = value
      End Set
    End Property

    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Property IsKey() As Boolean
      Get
        Return mblnIsKey
      End Get
      Set(ByVal value As Boolean)
        mblnIsKey = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpValue As Object)
      Me.New(lpName, Core.PropertyType.ecmString, lpValue, False)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpType As Core.PropertyType, ByVal lpValue As Object, ByVal lpIsKey As Boolean)
      Try
        menuType = lpType
        mstrName = lpName
        Value = lpValue
        mblnIsKey = lpIsKey
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpProperty As IProperty)
      Me.New(lpProperty, False)
    End Sub

    Public Sub New(lpProperty As IProperty, ByVal lpIsKey As Boolean)
      Try
        menuType = lpProperty.Type
        mstrName = lpProperty.Name
        Value = lpProperty.Value
        mblnIsKey = lpIsKey
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Overrides Function DebuggerIdentifier() As String
      Try
        If MyBase.Value IsNot Nothing Then
          Return String.Format("{0}: {1}", Name, Value.ToString)
        Else
          Return String.Format("{0}: No Value")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace