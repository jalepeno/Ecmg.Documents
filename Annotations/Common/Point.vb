'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Annotations.Decoration
Imports Documents.Utilities

#End Region

Namespace Annotations.Common

  ''' <summary>
  ''' Binds two related values together as coordinates of an unspecified two-dimensional geometry.
  ''' </summary>
  <Serializable()>
  Public Class Point

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Point"/> class.
    ''' </summary>
    Public Sub New()
      Try
        Me.LineEnd = New PointStyle()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Point"/> class.
    ''' </summary>
    ''' <param name="first">The first component of the coordinate.</param>
    ''' <param name="second">The second component of the coordinate.</param>
    Public Sub New(ByVal first As Single, ByVal second As Single)
      Try
        Me.First = first
        Me.Second = second
        Me.LineEnd = New PointStyle()
        'Me.Relative = False
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the line end.
    ''' </summary>
    ''' <value>The line end.</value>
    Public Property LineEnd() As PointStyle

    ''' <summary>
    ''' Gets or sets the first component of the coordinate.
    ''' </summary>
    ''' <value>The first component of the coordinate.</value>
    <XmlAttribute(AttributeName:="first")>
    Public Property First() As Single

    ''' <summary>
    ''' Gets or sets the second component of the coordinate.
    ''' </summary>
    ''' <value>The second component of the coordinate.</value>
    <XmlAttribute(AttributeName:="second")>
    Public Property Second() As Single

#End Region

  End Class

End Namespace