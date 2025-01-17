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
  ''' 
  ''' </summary>
  ''' <remarks>
  ''' 
  ''' Sample XML 
  ''' 
  ''' 		&lt;Action xsi:type="ChangePropertyValue" Name="Replace DocumentType"&gt;
  '''			&lt;PropertyName&gt;DocumentType&lt;/PropertyName&gt;
  '''			&lt;PropertyScope&gt;VersionProperty&lt;/PropertyScope&gt;
  '''			&lt;VersionIndex&gt;0&lt;/VersionIndex&gt;
  '''			&lt;SourceType&gt;DataLookup&lt;/SourceType&gt;
  '''			&lt;DataLookup xsi:type="DataList" CaseSensitive="false"&gt;
  '''				&lt;SourceProperty&gt;
  '''					&lt;PropertyName&gt;DocumentType&lt;/PropertyName&gt;
  '''					&lt;PropertyScope&gt;VersionProperty&lt;/PropertyScope&gt;
  '''					&lt;VersionIndex&gt;0&lt;/VersionIndex&gt;
  '''				&lt;/SourceProperty&gt;
  '''				&lt;DestinationProperty&gt;
  '''					&lt;PropertyName&gt;DocumentType&lt;/PropertyName&gt;
  '''					&lt;PropertyScope&gt;VersionProperty&lt;/PropertyScope&gt;
  '''					&lt;VersionIndex&gt;0&lt;/VersionIndex&gt;
  '''				&lt;/DestinationProperty&gt;
  '''				&lt;MapList&gt;
  '''					&lt;ValueMap Original="Argument" Replacement="Arguments"/&gt;
  '''					&lt;ValueMap Original="REPORT" Replacement="Report"/&gt;
  '''					&lt;ValueMap Original="DECISION.REPORT" Replacement="REPORT"/&gt;
  '''					&lt;ValueMap Original="Decision.Report" Replacement="Report"/&gt;
  '''					&lt;ValueMap Original="Interrogatories" Replacement="Interrogatory"/&gt;
  '''				&lt;/MapList&gt;
  '''			&lt;/DataLookup&gt;
  '''		&lt;/Action&gt;
  '''</remarks>
  <Serializable()>
  Public Class DataList
    Inherits DataLookup
    Implements IPropertyLookup

#Region "Class Variables"

    Private mobjSourceProperty As LookupProperty
    Private mobjDestinationProperty As LookupProperty
    Private mobjMapList As New MapList(Me)
    Private lblnCaseSensitive As Boolean

#End Region

#Region "Public Properties"

    <XmlAttribute()>
    Public Property CaseSensitive() As Boolean
      Get
        Return lblnCaseSensitive
      End Get
      Set(ByVal value As Boolean)
        lblnCaseSensitive = value
      End Set
    End Property

    ''' <summary>The property containing the source value.</summary>
    Public Property SourceProperty() As LookupProperty Implements IPropertyLookup.SourceProperty
      Get
        Return mobjSourceProperty
      End Get
      Set(ByVal Value As LookupProperty)
        mobjSourceProperty = Value
      End Set
    End Property

    ''' <summary>The property to update.</summary>
    Public Property DestinationProperty() As LookupProperty Implements IPropertyLookup.DestinationProperty
      Get
        Return mobjDestinationProperty
      End Get
      Set(ByVal Value As LookupProperty)
        mobjDestinationProperty = Value
      End Set
    End Property

    ''' <summary>The list of mapped values</summary>
    Public Property MapList() As MapList
      Get
        Return mobjMapList
      End Get
      Set(ByVal Value As MapList)
        mobjMapList = Value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(LookupType.List)
    End Sub

    Public Sub New(ByVal lpSourceProperty As LookupProperty,
                     ByVal lpDestinationProperty As LookupProperty)
      MyBase.New(LookupType.List)
      SourceProperty = lpSourceProperty
      DestinationProperty = lpDestinationProperty

    End Sub
#End Region

#Region "Public Methods"

    Public Overrides Function SourceExists(ByVal lpMetaHolder As Core.IMetaHolder) As Boolean
      Try
        If Me.SourceProperty IsNot Nothing AndAlso lpMetaHolder.Metadata.PropertyExists(SourceProperty.PropertyName, False) Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function
    Public Overrides Function GetParameters() As IParameters Implements IPropertyLookup.GetParameters
      Try

        Dim lobjParameters As IParameters = New Parameters

        If SourceProperty Is Nothing Then
          Throw New InvalidOperationException("The source property is not initialized.")
        End If

        If DestinationProperty Is Nothing Then
          Throw New InvalidOperationException("The destination property is not initialized.")
        End If

        For Each lobjLookupPropertyParameter As IParameter In SourceProperty.GetParameters()
          lobjLookupPropertyParameter.Name = String.Format("Source{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next
        For Each lobjLookupPropertyParameter As IParameter In DestinationProperty.GetParameters()
          lobjLookupPropertyParameter.Name = String.Format("Destination{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function GetValue(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Try

        Dim lobjSourceProperty As Core.ECMProperty
        Dim lstrSourcePropertyValue As String

        lobjSourceProperty = GetProperty(SourceProperty, lpMetaHolder)
        lstrSourcePropertyValue = lobjSourceProperty.Value

        Return MapList.FindReplacement(lstrSourcePropertyValue)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function GetValues(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Return Nothing
    End Function


#End Region

  End Class

End Namespace
