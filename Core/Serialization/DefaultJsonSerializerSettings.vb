' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DefaultJsonSerializerSettings.vb
'  Description :  [type_description_here]
'  Created     :  01/22/2024 10:49:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization

Namespace Core

  Public Class DefaultJsonSerializerSettings

#Region "Class Variables"

    Private mobjJsonSettings As New JsonSerializerSettings()
    Private mobjJsonSerializer As New JsonSerializer

#End Region

#Region "Constructors"

    Public Sub New()
      With mobjJsonSettings
        .TypeNameHandling = TypeNameHandling.Auto
        .ContractResolver = New CamelCasePropertyNamesContractResolver()
      End With
      With mobjJsonSerializer
        .TypeNameHandling = TypeNameHandling.Auto
        .ContractResolver = New CamelCasePropertyNamesContractResolver()
      End With
    End Sub

#End Region

#Region "Shared Properties"

    Public Shared ReadOnly Property Serializer As JsonSerializer
      Get
        Dim lobjSettings As New DefaultJsonSerializerSettings()
        Return lobjSettings.mobjJsonSerializer
      End Get
    End Property

    Public Shared ReadOnly Property Settings As JsonSerializerSettings

#End Region

  End Class

End Namespace