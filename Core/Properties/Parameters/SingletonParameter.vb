'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 1:54:13 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public MustInherit Class SingletonParameter
    Inherits Parameter
    Implements ISingletonProperty

#Region "Constructors"

    Protected Friend Sub New()
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpType As PropertyType)
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpType As PropertyType, ByVal lpName As String)
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
        Name = lpName
        SystemName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpType As PropertyType, ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Object)
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
        Name = lpName
        SystemName = lpSystemName
        Value = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Public Shadows Property Value As Object Implements ISingletonProperty.Value
      Get
        Try
          Return MyBase.Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Object)
        Try
          MyBase.Value = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

  End Class

End Namespace
