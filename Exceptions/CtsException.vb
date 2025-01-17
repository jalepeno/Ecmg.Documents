'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CtsException.vb
'   Description :  [type_description_here]
'   Created     :  1/13/2014 9:41:45 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Exceptions

  Public Class CtsException
    Inherits Exception

    Public Shared Function ExceptionToJson(lpException As Exception) As String
      Try
        Dim lobjStringWriter As StringWriter
        lobjStringWriter = New StringWriter
        Using lobjJsonWriter As New JsonTextWriter(lobjStringWriter)

          With lobjJsonWriter
            If Helper.IsRunningInstalled Then
              .Formatting = Formatting.None
            Else
              .Formatting = Formatting.Indented
            End If
            .WriteStartObject()
            .WritePropertyName("className")
            .WriteValue(lpException.GetType.Name)
            .WritePropertyName("message")
            .WriteValue(lpException.Message)
            If lpException.Data IsNot Nothing Then
              If TypeOf lpException.Data Is ICollection Then
                If lpException.Data.Count > 0 Then
                  .WritePropertyName("data")
                  .WriteValue(lpException.Data.ToString)
                End If
              End If
            End If
            .WriteEndObject()
          End With

          Return lobjStringWriter.ToString
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToJson() As String
      Try
        Return ExceptionToJson(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Overrides Function ToString() As String

      Dim lobjStringBuilder As New StringBuilder

      Try

        lobjStringBuilder.AppendFormat("{0},message:""{1}""", Me.GetType.Name, Message)

        If Data IsNot Nothing Then
          lobjStringBuilder.AppendFormat(",""data"":{0}", Data.ToString)
        End If

        If InnerException IsNot Nothing Then
          lobjStringBuilder.AppendFormat(",""innerException"":{0}", InnerException.ToString)
        End If

#If Not SILVERLIGHT = 1 Then
        If Not String.IsNullOrEmpty(Source) Then
          lobjStringBuilder.AppendFormat(",""source"":{0}", Source)
        End If
#End If

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return MyBase.ToString
      End Try
    End Function

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub
    Public Sub New(ByVal message As String, ByVal inner As Exception)
      MyBase.New(message, inner)
    End Sub

#If Not SILVERLIGHT = 1 Then
    Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
      MyBase.New(info, context)
    End Sub
#End If

#End Region

  End Class

End Namespace