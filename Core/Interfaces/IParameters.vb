' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IParameters.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 3:28:20 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


Imports System.ComponentModel

#End Region

Namespace Core

  Public Interface IParameters
    Inherits ICollection(Of IParameter)
    Inherits ICloneable

    Property Item(name As String) As IParameter
    Sub AddRange(lpParameters As IParameters)
    Function GetItemByIndex(ByVal index As Integer) As IParameter
    Sub SetItemByIndex(ByVal index As Integer, ByVal value As IParameter)
    Function GetItemByName(ByVal name As String) As IParameter
    Sub SetItemByName(ByVal name As String, ByVal value As IParameter)

    Overloads Function Contains(name As String) As Boolean
    Function GetValue(lpName As String) As Object
    Event ItemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)

  End Interface

End Namespace