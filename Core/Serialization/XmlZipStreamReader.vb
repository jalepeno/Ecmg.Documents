'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

'Imports System.IO.Compression
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace Serialization

  ''' <summary>
  ''' Represents a reader that provides fast, non-cached, forward only access to XML data from a stream.
  ''' </summary>
  ''' <remarks>Can also handle compressed xml data using the Ionic.Utils.Zip library.</remarks>
  Public Class XmlZipStreamReader
    Inherits XmlStreamReader

#Region "Class Variables"

    Private mobjZipFile As ZipFile

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the zip file that the stream was read from.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ZipFile() As ZipFile
      Get
        Return mobjZipFile
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Constructs a new XmlZipStreamReader using the specified stream and zip file.
    ''' </summary>
    ''' <param name="stream">Position should be set to zero</param>
    ''' <param name="zipFile">Zip file should already be open.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As IO.Stream, ByVal zipFile As ZipFile)
      MyBase.New(stream)
      Try
        mobjZipFile = zipFile
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

#End Region

  End Class

End Namespace