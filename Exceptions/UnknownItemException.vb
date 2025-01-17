'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  UnknownItemException.vb
'   Description :  [type_description_here]
'   Created     :  7/22/2013 10:44:18 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  Public Class UnknownItemException
    Inherits CtsException

#Region "Class Variables"

    Private mstrRequestedItem As String = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property RequestedItem As String
      Get
        Return mstrRequestedItem
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpRequestedItem As String)
      Me.New(String.Format(
             "Item '{0}' unknown.",
             lpRequestedItem), lpRequestedItem)
    End Sub

    Public Sub New(lpMessage As String, lpRequestedItem As String)
      MyBase.New(lpMessage)
      Try
        mstrRequestedItem = lpRequestedItem
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpMessage As String, lpRequestedItem As String, lpInnerException As Exception)
      MyBase.New(lpMessage, lpInnerException)
      Try
        mstrRequestedItem = lpRequestedItem
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace