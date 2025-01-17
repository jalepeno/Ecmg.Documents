' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ProviderInterfaceNotImplementedException.vb
'  Description :  [type_description_here]
'  Created     :  2/17/2012 2:25:43 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers
Imports Documents.Utilities

#End Region

Namespace Exceptions

  ''' <summary>
  ''' Exception thrown when an interface is requested for a provider that does not support it.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ProviderInterfaceNotImplementedException
    Inherits CtsException

#Region "Class Variables"

    Private mobjProvider As IProvider
    Private menuRequestedInterface As ProviderClass

#End Region

#Region "Public Properties"

    Public ReadOnly Property Provider As IProvider
      Get
        Return mobjProvider
      End Get
    End Property

    Public ReadOnly Property RequestedInterface As ProviderClass
      Get
        Return menuRequestedInterface
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpProvider As IProvider, lpRequestedInterface As ProviderClass)
      MyBase.New(String.Format("Requested interface '{0}' not implemented by provider '{1}'.",
                               lpRequestedInterface.ToString, lpProvider.Name))
      Try
        mobjProvider = lpProvider
        menuRequestedInterface = lpRequestedInterface
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpMessage As String, lpProvider As IProvider, lpRequestedInterface As ProviderClass)
      MyBase.New(lpMessage)
      Try
        mobjProvider = lpProvider
        menuRequestedInterface = lpRequestedInterface
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace