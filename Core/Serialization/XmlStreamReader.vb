'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports System.Globalization

#Region "Imports"

Imports System.Xml
Imports Documents.Utilities

#End Region

Namespace Serialization

  ''' <summary>
  ''' Represents a reader that provides fast, non-cached, forward only access to XML data from a stream.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class XmlStreamReader
    Inherits XmlTextReader

#Region "Class Variables"

    Private mobjStream As IO.Stream
    Private mstrLocaleString As String = String.Empty
    Private mobjCulture As CultureInfo

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the stream that was used to construct the reader.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Stream() As IO.Stream
      Get
        Return mobjStream
      End Get
    End Property

    Public ReadOnly Property LocaleString As String
      Get
        Return mstrLocaleString
      End Get
    End Property

    Public ReadOnly Property Culture As CultureInfo
      Get
        Return mobjCulture
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Constructs a new XmlStreamReader using the specified stream.
    ''' </summary>
    ''' <param name="stream">Position should be set to zero</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As IO.Stream)
      MyBase.New(stream)
      Try
        mobjStream = stream

#If CTSDOTNET = 1 Then
        GetLocale()
#End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

#If CTSDOTNET = 1 Then

    Private Sub GetLocale()
      Try
        Dim lintCurrentPosition As Integer = mobjStream.Position
        If mobjStream IsNot Nothing Then
          If mobjStream.CanSeek Then
            mobjStream.Position = 0
        End If
          mstrLocaleString = Helper.GetOriginalLocale(mobjStream)

          Try
            mobjCulture = New CultureInfo(mstrLocaleString, True)
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
            mobjCulture = CultureInfo.CurrentCulture
          End Try

          If mobjStream.CanSeek Then
            mobjStream.Position = lintCurrentPosition
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End If

#End Region

  End Class

End Namespace