'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonStringParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 1:54:49 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Xml.Serialization

#End Region

Namespace Core

  Public Class SingletonStringParameter
    Inherits SingletonParameter

#Region "Public Properties"

    <XmlIgnore()>
    Public Shadows Property Value As String
      Get
        Return MyBase.Value
      End Get
      Set(ByVal value As String)
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmString)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmString, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As String)
      MyBase.New(PropertyType.ecmString, lpName, lpSystemName, lpValue)
    End Sub

#End Region

  End Class

End Namespace