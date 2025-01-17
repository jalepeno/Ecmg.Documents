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

  ''' <summary>
  ''' A collection of ClassActionSet items for 
  ''' providing a menu of actions to perform 
  ''' for a variety of different document classes.
  ''' </summary>
  ''' <remarks></remarks>
  <XmlRoot("DocumentClassActionSets")>
  Public Class ClassActionsSets
    Inherits CCollection(Of ClassActionSet)

#Region "Class Constants"

    Public Property DEFAULT_IDENTIFIER As String = "Default"

#End Region

    Public Function ContainsClass(lpDocumentClassName As String) As Boolean
      Try

        If Item(lpDocumentClassName) IsNot Nothing Then
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Default Shadows Property Item(ByVal lpIndex As Integer) As ClassActionSet
      Get
        Try
          Return MyBase.Item(lpIndex)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As ClassActionSet)
        Try
          MyBase.Item(lpIndex) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the item for the specified document class name.  
    ''' </summary>
    ''' <param name="lpDocumentClassName">The document class to target.  
    ''' If the document class name is not specified, returns the default set.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Shadows Property Item(ByVal lpDocumentClassName As String) As ClassActionSet
      Get
        Try

          If String.IsNullOrEmpty(lpDocumentClassName) Then
            Return DefaultSet
          End If

          Dim lobjItem As ClassActionSet = Nothing
          lobjItem = Me.FirstOrDefault(Function(ActionSet) ActionSet.DocumentClass = lpDocumentClassName)

          Return lobjItem

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As ClassActionSet)
        Try
          Dim lobjItem As ClassActionSet = Nothing
          lobjItem = Me.FirstOrDefault(Function(ActionSet) ActionSet.DocumentClass = lpDocumentClassName)

          lobjItem = value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Property DefaultSet As ClassActionSet
      Get
        If Item(DEFAULT_IDENTIFIER) Is Nothing Then
          Add(New ClassActionSet(DEFAULT_IDENTIFIER))
        End If
        Return Item(DEFAULT_IDENTIFIER)
      End Get
      Set(value As ClassActionSet)
        Item(DEFAULT_IDENTIFIER) = value
      End Set
    End Property

    Public ReadOnly Property TargetClassCount As Integer
      Get
        Try
          Return Aggregate CAS In Me Where (CAS.DocumentClass <> DEFAULT_IDENTIFIER) Into Count()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
    End Property

  End Class

End Namespace