'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MultiValueParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:07:17 PM
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

  Public MustInherit Class MultiValueParameter
    Inherits Parameter
    Implements IMultiValuedProperty

#Region "Public Properties"

    Public Shadows Property Values As Values Implements IMultiValuedProperty.Values
      Get
        Return MyBase.Values
      End Get
      Set(ByVal value As Values)
        MyBase.Values = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpType As PropertyType)
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
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
        Cardinality = Core.Cardinality.ecmMultiValued
        Type = lpType
        Name = lpName
        SystemName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpType As PropertyType, ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(String.Empty)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
        Type = lpType
        Name = lpName
        SystemName = lpSystemName
        Values = lpValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace