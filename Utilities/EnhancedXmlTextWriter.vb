'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Globalization
Imports System.Xml

Namespace Utilities

  ''' <summary>
  ''' Enhances the XmlTextWriter class by exposing the file name
  ''' </summary>
  ''' <remarks></remarks>
  Public Class EnhancedXmlTextWriter
    Inherits XmlTextWriter

#Region "Class Variables"

    Private mstrFileName As String = String.Empty
    Private mobjCulture As CultureInfo = CultureInfo.CurrentCulture

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the file name that was supplied to the constructor
    ''' </summary>
    ''' <value>The fully qualified file name</value>
    ''' <returns></returns>
    ''' <remarks>If constructed with out a file name, this property will be empty</remarks>
    Public ReadOnly Property FileName() As String
      Get
        Return mstrFileName
      End Get
    End Property

    Public Property Culture As CultureInfo
      Get
        Return mobjCulture
      End Get
      Set(value As CultureInfo)
        mobjCulture = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal w As System.IO.TextWriter)
      MyBase.New(w)
    End Sub

    Public Sub New(ByVal w As System.IO.Stream, ByVal encoding As System.Text.Encoding)
      MyBase.New(w, encoding)
    End Sub

    Public Sub New(ByVal filename As String, ByVal encoding As System.Text.Encoding)
      MyBase.New(filename, encoding)
      mstrFileName = filename
    End Sub

#End Region

  End Class

End Namespace