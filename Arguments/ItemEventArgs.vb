' ********************************************************************************
' '  Document    :  ItemEventArgs.vb
' '  Description :  [type_description_here]
' '  Created     :  9/29/2012-12:35:39
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class ItemEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mobjItem As Object

#End Region

#Region "Public Properties"

    Public ReadOnly Property Item As Object
      Get
        Try
          Return mobjItem
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(item As Object)
      Try
        mobjItem = item
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace