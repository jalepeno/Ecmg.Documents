' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ParameterValueNotSetException.vb
'  Description :  [type_description_here]
'  Created     :  8/8/2012 3:40:37 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  Public Class ParameterValueNotSetException
    Inherits CtsException

#Region "Class Variables"

    Private mstrParameterName As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ParameterName As String
      Get
        Try
          Return mstrParameterName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpParameterName As String)
      MyBase.New(String.Format("Value not set for parameter '{0}'.", lpParameterName))
      Try
        mstrParameterName = lpParameterName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpParameterName As String, lpmessage As String)
      MyBase.New(lpmessage)
      Try
        mstrParameterName = lpParameterName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace