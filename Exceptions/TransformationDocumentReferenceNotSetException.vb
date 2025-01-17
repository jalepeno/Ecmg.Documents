'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

#End Region

''' <summary>
''' Contains the defined exceptions of the Cts framework
''' </summary>
Namespace Exceptions

  Public Class TransformationDocumentReferenceNotSetException
    Inherits DocumentReferenceNotSetException

#Region "Class Variables"

    Private mobjTransformation As Transformations.Transformation

#End Region

#Region "Public Properties"

    Public ReadOnly Property Transformation() As Transformations.Transformation
      Get
        Return mobjTransformation
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal transformation As Transformations.Transformation)
      MyBase.New()
      mobjTransformation = transformation
    End Sub

    Public Sub New(ByVal transformation As Transformations.Transformation, ByVal message As String)
      MyBase.New()
      mobjTransformation = transformation
    End Sub

#End Region

  End Class

End Namespace