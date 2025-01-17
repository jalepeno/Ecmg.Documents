'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IActionItems.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 2:11:22 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


Imports System.ComponentModel

#End Region

Namespace Core

  Public Interface IActionItems
    Inherits ICollection(Of IActionItem)
    Inherits ICloneable

    Property Item(name As String) As IActionItem
    Sub AddRange(lpItems As IActionItems)
    Function GetItemByIndex(ByVal index As Integer) As IActionItem
    Sub SetItemByIndex(ByVal index As Integer, ByVal value As IActionItem)
    Function GetItemByName(ByVal name As String) As IActionItem
    Sub SetItemByName(ByVal name As String, ByVal value As IActionItem)

    Overloads Function Contains(name As String) As Boolean
    'Function GetValue(lpName As String) As Object
    Event ItemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)

  End Interface

End Namespace