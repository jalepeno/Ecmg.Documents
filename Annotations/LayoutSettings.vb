'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports AnnotationsCommon = Documents.Annotations.Common

#End Region

Namespace Annotations

  ''' <summary>
  ''' Provides positioning, orientation and overall rectangular extent information.
  ''' </summary>
  <Serializable()>
  Public Class LayoutSettings

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the position.
    ''' </summary>
    ''' <value>The point of positioning.</value>
    Public Property Position() As AnnotationsCommon.Point

    ''' <summary>
    ''' Gets or sets the pivot point, or center for polar projection.
    ''' </summary>
    ''' <value>The pivot point.</value>
    Public Property Pivot() As AnnotationsCommon.Point

    ''' <summary>
    ''' Gets or sets the page number.
    ''' </summary>
    ''' <value>The page number.</value>
    <XmlAttribute()>
    Public Property PageNumber() As Integer

    ''' <summary>
    ''' Gets or sets the upper left extent.
    ''' </summary>
    ''' <value>The upper left extent.</value>
    Public Property UpperLeftExtent() As AnnotationsCommon.Point

    ''' <summary>
    ''' Gets or sets the lower right extent.
    ''' </summary>
    ''' <value>The lower right extent.</value>
    Public Property LowerRightExtent() As AnnotationsCommon.Point

    ''' <summary>
    ''' Gets or sets the angle.
    ''' </summary>
    ''' <value>The angle.</value>
    Public Property Angle() As Single

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Sets the extent.
    ''' </summary>
    ''' <param name="x1">The x1.</param>
    ''' <param name="y1">The y1.</param>
    ''' <param name="x2">The x2.</param>
    ''' <param name="y2">The y2.</param>
    Public Sub SetExtent(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

      ' -------------------------------------------------------------------------------
      ' Annotation layout
      ' -------------------------------------------------------------------------------
      ' Add the upper left and lower right points
      Dim upperLeft As New AnnotationsCommon.Point With {.First = x1, .Second = y1, .LineEnd = Nothing}
      Dim lowerRight As New AnnotationsCommon.Point With {.First = x2, .Second = y2, .LineEnd = Nothing}
      Me.UpperLeftExtent = upperLeft
      Me.LowerRightExtent = lowerRight
      Me.Position = upperLeft

    End Sub

    ''' <summary>
    ''' Updates the extent.
    ''' </summary>
    ''' <param name="x">The x coordinate.</param>
    ''' <param name="y">The y coordinate.</param>
    Public Sub UpdateExtent(ByVal x As Single, ByVal y As Single)

      If Me.UpperLeftExtent.First > x Then
        Me.UpperLeftExtent.First = x
      End If

      If Me.LowerRightExtent.First < x Then
        Me.LowerRightExtent.First = x
      End If

      If Me.UpperLeftExtent.Second > y Then
        Me.UpperLeftExtent.Second = y
      End If

      If Me.LowerRightExtent.Second < y Then
        Me.LowerRightExtent.Second = y
      End If

    End Sub

    Public Sub UpdatePivotAsCenter()
      Me.Pivot = New AnnotationsCommon.Point() With {
      .First = (Me.UpperLeftExtent.First + Me.LowerRightExtent.First) / 2,
      .Second = (Me.UpperLeftExtent.Second + Me.LowerRightExtent.Second) / 2,
      .LineEnd = Nothing
    }
    End Sub

#End Region

  End Class

End Namespace