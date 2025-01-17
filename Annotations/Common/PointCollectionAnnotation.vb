'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Xml.Serialization
Imports Documents.Annotations.Decoration
Imports Documents.Utilities


#End Region

Namespace Annotations.Common

  ''' <summary>
  ''' Annotation class describing lines, polygons, free-form lines and curves.
  ''' </summary>
  <Serializable()>
  <XmlInclude(GetType(ArrowAnnotation))>
  Public Class PointCollectionAnnotation
    Inherits LineBase

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PointCollectionAnnotation"/> class.
    ''' </summary>
    Public Sub New()
      Try
        Me.ConnectionType = PointConnectionType.Straight
        Me.LineStyle = New LineStyleInfo()
        Me.Points = New Collection(Of Point)()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' Describes the connection between points in the point collection, if any.
    ''' </summary>
    Public Enum PointConnectionType

      ''' <summary>
      ''' Indicates that the points are not connected.
      ''' </summary>
      None

      ''' <summary>
      ''' Indicates that the points are connected by straight lines.
      ''' </summary>
      Straight

      ''' <summary>
      ''' Indicates that a smooth curve passes through all of the points.
      ''' </summary>
      Spline
    End Enum

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the type of the connection.
    ''' </summary>
    ''' <value>The type of the connection.</value>
    Public Property ConnectionType() As PointConnectionType

    ''' <summary>
    ''' Gets or sets the points.
    ''' </summary>
    ''' <value>The points.</value>
    Public Property Points() As Collection(Of Point)

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Sets the starting point, clearing all other points from the collection.
    ''' </summary>
    ''' <param name="x">The x.</param>
    ''' <param name="y">The y.</param>
    ''' <param name="lineEnd">The line end.</param>
    Public Sub SetStartPoint(ByVal x As Single, ByVal y As Single, ByVal lineEnd As PointStyle)
      Me.Points.Clear()
      Dim startPoint As New Point With {.First = x, .Second = y, .LineEnd = lineEnd}
      Me.Layout.SetExtent(x, y, x, y)
      Me.Points.Add(startPoint)
      Me.Layout.UpdateExtent(x, y)
    End Sub

    ''' <summary>
    ''' Adds a line segment extending from the previous point.
    ''' </summary>
    ''' <param name="x">The x.</param>
    ''' <param name="y">The y.</param>
    ''' <param name="lineEnd">The line end.</param>
    Public Sub AddSegment(ByVal x As Single, ByVal y As Single, ByVal lineEnd As PointStyle)
      Dim startPoint As New Point With {.First = x, .Second = y, .LineEnd = lineEnd}
      Me.Points.Add(startPoint)
      Me.Layout.UpdateExtent(x, y)
    End Sub

#End Region

  End Class
End Namespace