'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Runtime.Serialization
Imports Documents.Exceptions

Namespace Annotations.Exception

  ''' <summary>
  ''' Marker class for all annotation exceptions.
  ''' </summary>
  Public MustInherit Class AnnotationException
    Inherits CtsException

    Public Sub New()

    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal serializationInfo As SerializationInfo, ByVal context As StreamingContext)
      MyBase.New(serializationInfo, context)
    End Sub

    Public Sub New(ByVal message As String, ByVal inner As System.Exception)
      MyBase.New(message, inner)
    End Sub

  End Class
End Namespace