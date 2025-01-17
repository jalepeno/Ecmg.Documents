'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Runtime.Serialization

#End Region

Namespace Annotations.Exception

  ''' <summary>
  ''' Thrown when a platform is asked to import an annotaiton that is neither native to the platform, nor can it reasonably construct.
  ''' </summary>
  Public Class UnsupportedAnnotationException
    Inherits AnnotationException

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="UnsupportedAnnotationException" /> class.
    ''' </summary>
    ''' <param name="serializationInfo">The serialization info.</param>
    ''' <param name="context">The context.</param>
    Public Sub New(ByVal serializationInfo As SerializationInfo, ByVal context As StreamingContext)
      MyBase.New(serializationInfo, context)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="UnsupportedAnnotationException" /> class.
    ''' </summary>
    ''' <param name="message">The message.</param>
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="UnsupportedAnnotationException" /> class.
    ''' </summary>
    ''' <param name="message">The message.</param>
    ''' <param name="inner">The inner.</param>
    Public Sub New(ByVal message As String, ByVal inner As System.Exception)
      MyBase.New(message, inner)
    End Sub

#End Region

  End Class
End Namespace