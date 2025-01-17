' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ConnectionStateChangedEventArgs.vb
'  Description :  To be fired when the connection state of a provider changes.
'  Created     :  10/17/2011 7:39:47 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the ConnectionStateChanged Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ConnectionStateChangedEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private menuCurrentState As Providers.ProviderConnectionState
    Private menuPreviousState As Nullable(Of Providers.ProviderConnectionState) = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property CurrentState() As Providers.ProviderConnectionState
      Get
        Return menuCurrentState
      End Get
    End Property

    Public ReadOnly Property PreviousState() As Nullable(Of Providers.ProviderConnectionState)
      Get
        Return menuPreviousState
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpCurrentState As Providers.ProviderConnectionState)
      MyBase.New()
      menuCurrentState = lpCurrentState
    End Sub

    Public Sub New(lpCurrentState As Providers.ProviderConnectionState, lpPreviousState As Providers.ProviderConnectionState)
      MyBase.New()
      menuCurrentState = lpCurrentState
      menuPreviousState = lpPreviousState
    End Sub

#End Region

  End Class

End Namespace
