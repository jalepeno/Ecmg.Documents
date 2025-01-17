'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  Document_Json.vb
'   Description :  [type_description_here]
'   Created     :  1/15/2014 12:26:01 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports Documents.Utilities
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization

#End Region

Namespace Core

  Partial Public Class Document

    Public Shared Function FromJson(lpJsonString As String) As Document
      Try
        Dim lobjJsonSettings As New JsonSerializerSettings()
        With lobjJsonSettings
          .TypeNameHandling = TypeNameHandling.Auto
          .ContractResolver = New CamelCasePropertyNamesContractResolver()
        End With

        Return JsonConvert.DeserializeObject(lpJsonString, GetType(Document), lobjJsonSettings)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToJson() As String
      Try

        If Helper.IsRunningInstalled Then
          Return JsonConvert.SerializeObject(Me, Formatting.None, DefaultJsonSerializerSettings.Settings)
        Else
          Return JsonConvert.SerializeObject(Me, Formatting.Indented, DefaultJsonSerializerSettings.Settings)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub ToJsonFile(lpFilePath As String)
      Try

        Dim lobjOutputStream As New MemoryStream
        Dim lobjJsonSerializer As New JsonSerializer

        With lobjJsonSerializer
          .TypeNameHandling = TypeNameHandling.Auto
          .ContractResolver = New CamelCasePropertyNamesContractResolver()
        End With

        Using lobjStreamWriter As New StreamWriter(lpFilePath)
          Using lobjJsonWriter As New JsonTextWriter(lobjStreamWriter)
            lobjJsonSerializer.Serialize(lobjJsonWriter, Me)
          End Using
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToJsonStream() As IO.Stream
      Try

        Return Helper.CopyStringToStream(Me.ToJson())

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "ShouldSerialize Methods"

    Public Function ShouldSerializeAuditEvents() As Boolean
      Try
        Return AuditEvents.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ShouldSerializeRelationships() As Boolean
      Try
        Return Relationships.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ShouldSerializePermissions() As Boolean
      Try
        Return Permissions.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ShouldSerializeRelatedDocuments() As Boolean
      Try
        Return RelatedDocuments.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace