'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ValueMap
    Implements IDescription

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrOriginal As String = String.Empty
    Private mstrReplacement As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the name of the object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property Name() As String Implements IDescription.Name, INamedItem.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the description of the object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property Description() As String Implements IDescription.Description
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the original value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This property is used as the criterion for the ValueMap</remarks>
    <XmlAttribute()>
    Public Property Original() As String
      Get
        Return mstrOriginal
      End Get
      Set(ByVal value As String)
        Try

          ' Make sure we get a valid string
          If value Is Nothing Then
            Throw New ArgumentNullException
          End If

          mstrOriginal = value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the replacement value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This property is used as the replacement when the original is found</remarks>
    <XmlAttribute()>
    Public Property Replacement() As String
      Get
        Return mstrReplacement
      End Get
      Set(ByVal value As String)
        Try

          ' Make sure we get a valid string
          If value Is Nothing Then
            Throw New ArgumentNullException
          End If

          mstrReplacement = value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal original As String, ByVal replacement As String)
      Me.Original = original
      Me.Replacement = replacement
    End Sub

#End Region

#Region "Public Methods"

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New Text.StringBuilder
      Try

        If Name IsNot Nothing OrElse Name.Length > 0 Then
          lobjIdentifierBuilder.AppendFormat("({0}) ", Name)
        End If

        lobjIdentifierBuilder.AppendFormat("Original:{0}, Replacement:{1}", Original, Replacement)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
