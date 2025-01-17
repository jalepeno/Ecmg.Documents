'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Arguments

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TransformDocumentEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mobjTransformation As Transformation

#End Region

#Region "Public Properties"

    Public Property Transformation As Transformation
      Get
        Return mobjTransformation
      End Get
      Set(ByVal value As Transformation)
        mobjTransformation = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document,
                   ByVal lpTransformation As Transformation)
      MyBase.New(lpDocument)
      Transformation = lpTransformation
    End Sub

    Public Sub New(ByVal lpDocument As Document,
                   ByVal lpTransformation As Transformation,
                   ByVal lpEventDescription As String)
      MyBase.New(lpDocument, lpEventDescription, Now)
      Transformation = lpTransformation
    End Sub

    Public Sub New(ByVal lpDocument As Document,
                   ByVal lpTransformation As Transformation,
                   ByVal lpTime As DateTime)
      MyBase.New(lpDocument, "", lpTime)
      Transformation = lpTransformation
    End Sub

    Public Sub New(ByVal lpDocument As Document,
                   ByVal lpTransformation As Transformation,
                   ByVal lpEventDescription As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpEventDescription, lpTime)
      Transformation = lpTransformation
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        Return String.Format("Transform={0}; {1}", Transformation.DebuggerIdentifier, MyBase.DebuggerIdentifier)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace