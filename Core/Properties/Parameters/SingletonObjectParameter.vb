'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonObjectParameter.vb
'   Description :  [type_description_here]
'   Created     :  5/18/2015 2:55:37 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Core

  Public Class SingletonObjectParameter
    Inherits SingletonParameter

#Region "Public Properties"

    '<XmlIgnore()> _
    Public Shadows Property Value As Object
      Get
        Return MyBase.Value
      End Get
      Set(ByVal value As Object)
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmObject)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmObject, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Object)
      MyBase.New(PropertyType.ecmObject, lpName, lpSystemName, lpValue)
    End Sub

#End Region

  End Class

End Namespace