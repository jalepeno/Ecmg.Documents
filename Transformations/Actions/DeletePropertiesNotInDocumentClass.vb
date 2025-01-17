'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

Namespace Transformations
  ''' <summary>
  ''' Eliminates properties not listed in destination document class
  ''' </summary>
  ''' <remarks>
  ''' Only works in cases where the import provider implements the 'IClassification' interface.
  ''' </remarks>
  <Serializable()>
  Public Class DeletePropertiesNotInDocumentClass
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "DeletePropertiesNotInDocumentClass"

#End Region

#Region "Class Variables"

    Private mobjDocumentClasses As DocumentClasses
    Private mstrDocumentClassName As String

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property DocumentClasses() As DocumentClasses
      Get
        Return mobjDocumentClasses
      End Get
      Set(ByVal value As DocumentClasses)
        mobjDocumentClasses = value
      End Set
    End Property

    <XmlAttribute("ClassName")>
    Public Property DocumentClassName() As String
      Get
        Return mstrDocumentClassName
      End Get
      Set(ByVal value As String)
        mstrDocumentClassName = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.DeletePropertiesNotInDocumentClass)
    End Sub

    Public Sub New(ByVal lpDocumentClasses As DocumentClasses, Optional ByVal lpPropertyScope As PropertyScope = PropertyScope.VersionProperty)
      MyBase.New(ActionType.DeletePropertiesNotInDocumentClass, lpPropertyScope)
      DocumentClasses = lpDocumentClasses
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      Try
        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If
      Catch ex As Exception
        lpErrorMessage = ex.Message
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try

      ' Check to make sure the document class is set
      If DocumentClasses Is Nothing Then
        Return New ActionResult(Me, False, "DocumentClasses variable not set")
      End If

      ' Save the names of the properties we delete
      Dim lstrDeletedProperties As String = "Deleted Properties: "
      Dim lstrDeletedPropertyDelimiter As String = ","

      Try

        ' Make sure the DocumentClass exists in the collection
        Dim lstrDocumentClassName As String
        lstrDocumentClassName = Me.Transformation.Document.DocumentClass '.Properties("Document Class").Value
        If lstrDocumentClassName.Length = 0 Then
          lpErrorMessage = "DocumentClass not specified"
          Return New ActionResult(Me, False, lpErrorMessage)
        End If
        If DocumentClasses.ClassExists(Me.Transformation.Document.DocumentClass) = False Then
          lpErrorMessage = "DocumentClass '" & Me.Transformation.Document.DocumentClass & "' does not exist in the DocumentClasses collection"
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Dim lobjDocumentClass As DocumentClass = DocumentClasses(lstrDocumentClassName)
        Dim lobjDeletionProperties As InvalidProperties = lobjDocumentClass.FindInvalidProperties(Me.Transformation.Document, Me.PropertyScope, lpErrorMessage)

        lstrDeletedProperties = DeleteProperties(lobjDeletionProperties)

        Return New ActionResult(Me, True, lstrDeletedProperties)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        If lpErrorMessage.Length = 0 Then
          lpErrorMessage = lstrDeletedProperties & vbCrLf & vbCrLf & ex.Message
        End If
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try

    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        ' Do nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Function DeleteProperties(ByVal lpDeletionProperties As InvalidProperties) As String

      Try

        ' Save the names of the properties we delete
        Dim lstrDeletedProperties As String = "Deleted Properties: "
        Dim lstrDeletedPropertyDelimiter As String = ","

        Select Case Me.PropertyScope
          Case PropertyScope.DocumentProperty
            For Each lobjDeletionProperty As InvalidProperty In lpDeletionProperties
              If Me.Transformation.Document.Properties.PropertyExists(lobjDeletionProperty.Name) Then
                Me.Transformation.Document.RemoveProperty(lobjDeletionProperty.BaseProperty)
              End If
            Next

          Case PropertyScope.VersionProperty
            For Each lobjVersion As Version In Me.Transformation.Document.Versions
              For Each lobjDeletionProperty As InvalidProperty In lpDeletionProperties
                If lobjVersion.Properties.PropertyExists(lobjDeletionProperty.Name) Then
                  lobjVersion.Properties.Remove(lobjDeletionProperty.BaseProperty)
                End If
              Next
            Next

          Case PropertyScope.BothDocumentAndVersionProperties
            For Each lobjDeletionProperty As InvalidProperty In lpDeletionProperties
              If Me.Transformation.Document.Properties.PropertyExists(lobjDeletionProperty.Name) Then
                Me.Transformation.Document.RemoveProperty(lobjDeletionProperty.BaseProperty)
              End If
            Next

            For Each lobjVersion As Version In Me.Transformation.Document.Versions
              For Each lobjDeletionProperty As InvalidProperty In lpDeletionProperties
                If lobjVersion.Properties.PropertyExists(lobjDeletionProperty.Name) Then
                  lobjVersion.Properties.Remove(lobjDeletionProperty.BaseProperty)
                End If
              Next
            Next

        End Select

        If lstrDeletedProperties.EndsWith(lstrDeletedPropertyDelimiter) Then
          lstrDeletedProperties = lstrDeletedProperties.Remove(lstrDeletedProperties.Length - lstrDeletedPropertyDelimiter.Length)
        End If

        Return lstrDeletedProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace