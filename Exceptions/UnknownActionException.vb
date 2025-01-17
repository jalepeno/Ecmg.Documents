'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  UnknownActionException.vb
'   Description :  [type_description_here]
'   Created     :  7/22/2013 10:51:32 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Exceptions

  Public Class UnknownActionException
    Inherits UnknownItemException

#Region "Public Properties"

    Public ReadOnly Property RequestedAction As String
      Get
        Return MyBase.RequestedItem
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpRequestedAction As String)
      'Me.New(String.Format( _
      '       "Action '{0}' unknown, if this is an extension operation make sure the extension is registered in ExtensionCatalog.xml.", _
      '       lpRequestedAction), lpRequestedAction)
      Me.New(String.Format(
         "Action '{0}' unknown.",
         lpRequestedAction), lpRequestedAction)
    End Sub

    Public Sub New(lpMessage As String, lpRequestedAction As String)
      MyBase.New(lpMessage, lpRequestedAction)
    End Sub

#End Region

  End Class

End Namespace