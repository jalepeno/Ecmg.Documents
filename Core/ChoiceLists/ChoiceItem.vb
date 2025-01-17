'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization


Namespace Core.ChoiceLists

  ''' <summary>
  ''' An individual item in a choice list
  ''' </summary>
  ''' <remarks></remarks>
  <XmlInclude(GetType(ChoiceGroup)),
  XmlInclude(GetType(ChoiceValue))>
  Public Class ChoiceItem
    Inherits RepositoryObject

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Name As String)
      MyBase.New(Name)
    End Sub

#End Region

  End Class

End Namespace