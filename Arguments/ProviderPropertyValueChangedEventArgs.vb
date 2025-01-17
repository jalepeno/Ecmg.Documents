'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the ProviderPropertyValueChanged Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ProviderPropertyValueChangedEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mobjProviderProperty As Providers.ProviderProperty
    Private mobjOldValue As Object
    Private mobjNewValue As Object

#End Region

#Region "Public Properties"

    Public ReadOnly Property ProviderProperty() As Providers.ProviderProperty
      Get
        Return mobjProviderProperty
      End Get
    End Property

    Public ReadOnly Property OldValue() As Object
      Get
        Return mobjOldValue
      End Get
    End Property

    Public ReadOnly Property NewValue() As Object
      Get
        Return mobjNewValue
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpProviderProperty As Providers.ProviderProperty, ByVal lpOldValue As Object, ByVal lpNewValue As Object)
      MyBase.New()
      mobjProviderProperty = lpProviderProperty
      mobjOldValue = lpOldValue
      mobjNewValue = lpNewValue
    End Sub

#End Region

  End Class
End Namespace