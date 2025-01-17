'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  NotifyObject.vb
'   Description :  [type_description_here]
'   Created     :  12/31/2013 7:43:44 AM
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

  Public Class NotifyObject
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub OnPropertyChanged(ByVal sProp As String)
      RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sProp))
    End Sub

  End Class

End Namespace