'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO

#End Region

Namespace Utilities

  ''' <summary>
  ''' Enhances the XmlTextWriter class by exposing the file name
  ''' </summary>
  ''' <remarks></remarks>
  Public Class EnhancedStreamReader
    Inherits StreamReader

#Region "Class Variables"

    Private _path As String = String.Empty

#End Region

#Region "Pubilc Properties"

    ''' <summary>
    ''' Gets the complete path to be read
    ''' </summary>
    ''' <value>The fully qualified file name</value>
    ''' <returns></returns>
    ''' <remarks>If constructed with out a path, this property will be empty</remarks>
    Public ReadOnly Property Path() As String
      Get
        Return _path
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader
    ''' class for the specified file name.
    ''' </summary>
    ''' <param name="path">The complete file path to be read.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">path is an empty string ("").</exception>
    ''' <exception cref="System.ArgumentNullException">path is null.</exception>
    ''' <exception cref="System.IO.FileNotFoundException">The file cannot be found.</exception>
    ''' <exception cref="System.IO.DirectoryNotFoundException">The specified path is invalid, 
    ''' such as being on an unmapped drive.</exception>
    ''' <exception cref="System.IO.IOException">path includes an incorrect or invalid syntax 
    ''' for file name, directory name, or volume label.</exception>
    Public Sub New(ByVal path As String)
      MyBase.New(path)
      _path = path
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader 
    ''' class for the specified file name, with the specified byte 
    ''' order mark detection option.
    ''' </summary>
    ''' <param name="path">The complete file path to be read.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether 
    ''' to look for byte order marks at the beginning of the file.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">path is an empty string ("").</exception>
    ''' <exception cref="System.ArgumentNullException">path is null.</exception>
    ''' <exception cref="System.IO.FileNotFoundException">The file cannot be found.</exception>
    ''' <exception cref="System.IO.DirectoryNotFoundException">The specified path is invalid, 
    ''' such as being on an unmapped drive.</exception>
    ''' <exception cref="System.IO.IOException">path includes an incorrect or invalid syntax 
    ''' for file name, directory name, or volume label.</exception>
    Public Sub New(ByVal path As String,
                 ByVal detectEncodingFromByteOrderMarks As Boolean)
      MyBase.New(path, detectEncodingFromByteOrderMarks)
      _path = path
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader 
    ''' class for the specified file name, 
    ''' with the specified character encoding.
    ''' </summary>
    ''' <param name="path">The complete file path to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">path is an empty string ("").</exception>
    ''' <exception cref="System.ArgumentNullException">path is null.</exception>
    ''' <exception cref="System.IO.FileNotFoundException">The file cannot be found.</exception>
    ''' <exception cref="System.IO.DirectoryNotFoundException">The specified path is invalid, 
    ''' such as being on an unmapped drive.</exception>
    ''' <exception cref="System.NotSupportedException">path includes an incorrect or invalid syntax 
    ''' for file name, directory name, or volume label.</exception>
    Public Sub New(ByVal path As String,
                 ByVal encoding As System.Text.Encoding)
      MyBase.New(path, encoding)
      _path = path
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader 
    ''' class for the specified file name, with the specified 
    ''' character encoding and byte order mark detection option.
    ''' </summary>
    ''' <param name="path">The complete file path to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">path is an empty string ("").</exception>
    ''' <exception cref="System.ArgumentNullException">path or encoding is null.</exception>
    ''' <exception cref="System.IO.FileNotFoundException">The file cannot be found.</exception>
    ''' <exception cref="System.IO.DirectoryNotFoundException">The specified path is invalid, 
    ''' such as being on an unmapped drive.</exception>
    ''' <exception cref="System.NotSupportedException">path includes an incorrect or invalid syntax 
    ''' for file name, directory name, or volume label.</exception>
    Public Sub New(ByVal path As String,
                 ByVal encoding As System.Text.Encoding,
                 ByVal detectEncodingFromByteOrderMarks As Boolean)
      MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
      _path = path
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader 
    ''' class for the specified file name, with the specified 
    ''' character encoding, byte order mark detection option, 
    ''' and buffer size.
    ''' </summary>
    ''' <param name="path">The complete file path to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
    ''' <param name="bufferSize">The minimum buffer size, in number of 16-bit characters.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">path is an empty string ("").</exception>
    ''' <exception cref="System.ArgumentNullException">path or encoding is null.</exception>
    ''' <exception cref="System.IO.FileNotFoundException">The file cannot be found.</exception>
    ''' <exception cref="System.IO.DirectoryNotFoundException">The specified path is invalid, 
    ''' such as being on an unmapped drive.</exception>
    ''' <exception cref="System.NotSupportedException">path includes an incorrect or invalid syntax 
    ''' for file name, directory name, or volume label.</exception>
    ''' <exception cref="System.ArgumentOutOfRangeException">buffersize is less than or equal to zero.</exception>
    Public Sub New(ByVal path As String,
                 ByVal encoding As System.Text.Encoding,
                 ByVal detectEncodingFromByteOrderMarks As Boolean,
                 ByVal bufferSize As Integer)
      MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
      _path = path
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader class for the specified stream.</summary>
    ''' <param name="stream">The stream to be read.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">The stream does not support reading.</exception>
    ''' <exception cref="System.ArgumentNullException">stream is null.</exception>
    Public Sub New(ByVal stream As System.IO.Stream)
      MyBase.New(stream)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader class for the specified stream, with the specified byte order mark detection option.</summary>
    ''' <param name="stream">The stream to be read.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">The stream does not support reading.</exception>
    ''' <exception cref="System.ArgumentNullException">stream is null.</exception>
    Public Sub New(ByVal stream As System.IO.Stream,
                  ByVal detectEncodingFromByteOrderMarks As Boolean)
      MyBase.New(stream, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader class for the specified stream, with the specified character encoding.</summary>
    ''' <param name="stream">The stream to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">The stream does not support reading.</exception>
    ''' <exception cref="System.ArgumentNullException">stream or encoding is null.</exception>
    Public Sub New(ByVal stream As System.IO.Stream,
                  ByVal encoding As System.Text.Encoding)
      MyBase.New(stream, encoding)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader class for the specified stream, with the specified character encoding and byte order mark detection option.</summary>
    ''' <param name="stream">The stream to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">The stream does not support reading.</exception>
    ''' <exception cref="System.ArgumentNullException">stream or encoding is null.</exception>
    Public Sub New(ByVal stream As System.IO.Stream,
                  ByVal encoding As System.Text.Encoding,
                  ByVal detectEncodingFromByteOrderMarks As Boolean)
      MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the EnhancedStreamReader class for the specified stream, with the specified character encoding, byte order mark detection option, and buffer size.</summary>
    ''' <param name="stream">The stream to be read.</param>
    ''' <param name="encoding">The character encoding to use.</param>
    ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
    ''' <param name="bufferSize">The minimum buffer size.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">The stream does not support reading.</exception>
    ''' <exception cref="System.ArgumentNullException">stream or encoding is null.</exception>
    ''' <exception cref="System.ArgumentOutOfRangeException">bufferSize is less than or equal to zero.</exception>
    Public Sub New(ByVal stream As System.IO.Stream,
                 ByVal encoding As System.Text.Encoding,
                 ByVal detectEncodingFromByteOrderMarks As Boolean,
                 ByVal bufferSize As Integer)
      MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
    End Sub

#End Region

  End Class

End Namespace