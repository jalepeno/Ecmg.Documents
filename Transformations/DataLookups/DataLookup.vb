'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Data
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations
  ''' <summary>Base class from which all other data lookup classes inherit.</summary>
  ''' <remarks>Used for changing a document property value by looking up the new value.</remarks>
  <XmlInclude(GetType(DataSource)),
    XmlInclude(GetType(DataParser)),
    XmlInclude(GetType(DataList))>
  <Serializable()>
  Public MustInherit Class DataLookup
    'Friend menuType As LookupType
    ''' <summary>Gets the value to apply to the destination property.</summary>
    Public MustOverride Function GetValue(ByVal lpMetaHolder As Core.IMetaHolder) As Object
    Public MustOverride Function GetValues(ByVal lpMetaHolder As Core.IMetaHolder) As Object
    Public MustOverride Function SourceExists(ByVal lpMetaHolder As Core.IMetaHolder) As Boolean
    'Public MustOverride Function GetValue(ByVal lpDocument As Core.Document) As Object
    'Public MustOverride Function GetValue(ByVal lpVersion As Core.Version) As Object

    Public MustOverride Function GetParameters() As IParameters

#Region "Public Properties"

    Friend Property Action As Action

    Public ReadOnly Property Name As String
      Get
        Try
          Dim lobjPropertyLookup As IPropertyLookup
          Dim lstrValueSourceType As String = String.Empty
          Dim lstrSourceValue As String = String.Empty




          If TypeOf (Me) Is DataMap Then
            lstrValueSourceType = String.Format("DataMap(from {0})",
                                                CType(Me, DataMap).SourceColumn)
            'ElseIf (TypeOf (Me) Is DataParser) OrElse (TypeOf (Me) Is DataList) Then
          ElseIf (TypeOf (Me) Is IPropertyLookup) Then
            lobjPropertyLookup = CType(Me, IPropertyLookup)
            If Action IsNot Nothing AndAlso Action.Transformation IsNot Nothing AndAlso Action.Transformation.Document IsNot Nothing Then
              Try
                If Action.Transformation.Document.PropertyExists(lobjPropertyLookup.SourceProperty.PropertyScope, lobjPropertyLookup.SourceProperty.PropertyName, False) Then
                  lstrSourceValue = Action.Transformation.Document.GetProperty(lobjPropertyLookup.SourceProperty.PropertyName, lobjPropertyLookup.SourceProperty.VersionIndex).Value.ToString
                End If
              Catch ex As Exception
                lstrSourceValue = String.Empty
              End Try
            End If

            If TypeOf (Me) Is DataParser Then

              lstrValueSourceType = String.Format("DataParser({0} from {1}({2}))",
                                                  lobjPropertyLookup.DestinationProperty.Name, lobjPropertyLookup.SourceProperty.Name, lstrSourceValue)
            ElseIf TypeOf (Me) Is DataList Then
              lstrValueSourceType = String.Format("DataList({0} from {1}({2}))",
                                                  lobjPropertyLookup.DestinationProperty.Name, lobjPropertyLookup.SourceProperty.Name, lstrSourceValue)
            End If
          End If
          If lstrValueSourceType.Length > 0 Then
            Return lstrValueSourceType
          Else
            Return Me.GetType.Name
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '' Re-throw the exception to the caller
          'Throw
          Return Me.GetType.Name
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpLookupType As LookupType)

    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Gets the specified property
    ''' </summary>
    ''' <param name="lpLookupProperty">The property to return</param>
    ''' <param name="lpAction">The current transformation action</param>
    ''' <returns>An ECMProperty object reference</returns>
    ''' <remarks>Gets the property from the document associated 
    ''' with the parent transformation of the action parameter</remarks>
    ''' <exception cref="TransformationDocumentReferenceNotSetException">
    ''' If the document reference is not set in the transformation associated with the action parameter, 
    ''' a TransformationDocumentReferenceNotSetException will be thrown.
    ''' </exception>
    ''' <exception cref="PropertyDoesNotExistException">
    ''' If the property does not exist, a PropertyDoesNotExistException will be thrown.
    ''' </exception>
    Public Shared Function GetProperty(ByVal lpLookupProperty As LookupProperty, ByVal lpAction As Action) As Core.ECMProperty
      Try

        ' Get the document first
        Dim lobjDocument As Document = lpAction.Transformation.Document
        Dim lobjProperty As Core.ECMProperty = TryGetProperty(lpLookupProperty, lpAction)

        If lobjProperty IsNot Nothing Then
          Return lobjProperty
        ElseIf lpLookupProperty.AutoCreate = True Then
          ' Create the property and return it
          If lobjDocument.CreateProperty(lpLookupProperty.PropertyName, , PropertyType.ecmString, lpLookupProperty.PropertyScope) = True Then
            Return lobjDocument.GetProperty(lpLookupProperty)
          Else
            Throw New PropertyDoesNotExistException(
              String.Format("The property '{0}' could not be created.",
                            lpLookupProperty.ToString),
              lpLookupProperty.PropertyName)
          End If
        Else
          Throw New PropertyDoesNotExistException(
            String.Format("The lookup property '{0}' could not be found.",
                          lpLookupProperty.ToString),
            lpLookupProperty.PropertyName)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Trys to get the specified property
    ''' </summary>
    ''' <param name="lpLookupProperty">The property to return</param>
    ''' <param name="lpAction">The current transformation action</param>
    ''' <returns>An ECMProperty object reference if th eproperty exists, otherwise returns Nothing</returns>
    ''' <remarks>Gets the property from the document associated 
    ''' with the parent transformation of the action parameter
    ''' </remarks>
    Public Shared Function TryGetProperty(ByVal lpLookupProperty As LookupProperty, ByVal lpAction As Action) As Core.ECMProperty
      Try

        ' Get the document first
        Dim lobjDocument As Document = lpAction.Transformation.Document
        Dim lobjProperty As Core.ECMProperty = Nothing

        ' Find out where we need to get the property from
        Select Case lpLookupProperty.PropertyScope
          Case PropertyScope.DocumentProperty, PropertyScope.BothDocumentAndVersionProperties, PropertyScope.AllProperties
            ' We will get the property from the set of document properties
            lobjProperty = TryGetProperty(lpLookupProperty, lobjDocument)
          Case PropertyScope.VersionProperty
            ' We will get the property from the set of version properties
            lobjProperty = TryGetProperty(lpLookupProperty, lobjDocument.Versions(lpLookupProperty.VersionIndex))
          Case PropertyScope.ContentProperty
            ' We will get the property from the set of content properties
            lobjProperty = TryGetProperty(lpLookupProperty, lobjDocument.Versions(lpLookupProperty.VersionIndex).GetPrimaryContent)
        End Select

        Return lobjProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets the specified property
    ''' </summary>
    ''' <param name="lpLookupProperty">The property to return</param>
    ''' <param name="lpMetaHolder">The object from which to retrieve the property</param>
    ''' <returns>An ECMProperty object reference</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentException">If the MetaHolder is not a Document, Version or Content object, an ArgumentException will be thrown.</exception>
    ''' <exception cref="PropertyDoesNotExistException">If the property does not exist, a PropertyDoesNotExistException exception will be thrown.</exception>
    Public Shared Function GetProperty(ByVal lpLookupProperty As LookupProperty, ByVal lpMetaHolder As Core.IMetaHolder) As Core.ECMProperty
      Try

        Dim lobjProperty As Core.ECMProperty = TryGetProperty(lpLookupProperty, lpMetaHolder)

        If lobjProperty IsNot Nothing Then
          Return lobjProperty
        Else
          ' The property was not found in the document, version or content
          Throw New PropertyDoesNotExistException(
            String.Format("The lookup property '{0}' could not be found in the {1} object",
                          lpLookupProperty.ToString,
                          lpMetaHolder.GetType.Name),
            lpLookupProperty.PropertyName)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Trys to get the specified property
    ''' </summary>
    ''' <param name="lpLookupProperty">The property to return</param>
    ''' <param name="lpMetaHolder">The object from which to retrieve the property</param>
    ''' <returns>An ECMProperty object reference if the property exists, otherwise returns Nothing.</returns>
    ''' <remarks></remarks>
    Public Shared Function TryGetProperty(ByVal lpLookupProperty As LookupProperty, ByVal lpMetaHolder As Core.IMetaHolder) As Core.ECMProperty
      Try

        Dim lobjProperty As Core.ECMProperty = Nothing

        If lpLookupProperty Is Nothing Then
          Throw New ArgumentNullException("lpLookupProperty", "Unable to get property, the parameter lpLookupProperty is null.")
        End If

        If lpMetaHolder Is Nothing Then
          Throw New ArgumentNullException("lpMetaHolder", "Unable to get property, the parameter lpMetaHolder is null.")
        End If

        Select Case lpMetaHolder.GetType.Name
          Case "Document"
            lobjProperty = CType(lpMetaHolder, Core.Document).GetProperty(lpLookupProperty.PropertyName, lpLookupProperty.PropertyScope, lpLookupProperty.VersionIndex)

          Case "Record"
            lobjProperty = CType(lpMetaHolder, Core.Document).GetProperty(lpLookupProperty.PropertyName, lpLookupProperty.PropertyScope, lpLookupProperty.VersionIndex)

          Case "Version"
            lobjProperty = lpMetaHolder.Metadata.Item(lpLookupProperty.PropertyName)

          Case "Content"
            lobjProperty = CType(lpMetaHolder, Content).CompleteProperties(lpLookupProperty.PropertyName)

          Case Else
            Throw New ArgumentException("The argument is not a Document, Version or Content object", "lpMetaHolder")
        End Select

        If lobjProperty IsNot Nothing Then
          Return lobjProperty
        Else
          ' The property was not found in the document, version or content
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace