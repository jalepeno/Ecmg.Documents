'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Comparison

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ComparisonProperty
    Implements IProperty
    Implements INotifyPropertyChanged
    Implements IXmlSerializable

#Region "Public Events"

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Public Enumerations"

    Public Enum MatchStatusEnum
      DoesNotExistInEitherDocument
      ExistsInBothDocuments
      DocumentXOnly
      DocumentYOnly
    End Enum

#End Region

#Region "Class Variables"

    Private mstrComparisonValue As String = Nothing
    Private menuMatchStatus As MatchStatusEnum = MatchStatusEnum.DoesNotExistInEitherDocument
    Private mobjPropertyX As IProperty = Nothing
    Private mobjPropertyY As IProperty = Nothing
    Private mstrXsiType As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property IsEqual As Boolean
      Get
        Try
          Return CompareProperties()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property ComparisonValue As String
      Get
        If mstrComparisonValue Is Nothing Then
          mstrComparisonValue = GetComparisonValue()
        End If
        Return mstrComparisonValue
      End Get
    End Property

    Public ReadOnly Property MatchStatus As MatchStatusEnum
      Get
        Return menuMatchStatus
      End Get
    End Property

    Public Overridable ReadOnly Property HasStandardValues As Boolean Implements IProperty.HasStandardValues
      Get
        Try
          Return False
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overridable ReadOnly Property StandardValues As IEnumerable Implements IProperty.StandardValues
      Get
        Try
          Return New List(Of String)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property PropertyX As IProperty
      Get
        Return mobjPropertyX
      End Get
      Set(value As IProperty)
        Try
          mobjPropertyX = value
          If value IsNot Nothing Then
            ' Sync up the IProperty values using PropertyX
            Helper.AssignObjectProperties(value, Me)
            RefreshMatchStatus()
            OnPropertyChanged("PropertyX")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property PropertyY As IProperty
      Get
        Return mobjPropertyY
      End Get
      Set(value As IProperty)
        Try
          mobjPropertyY = value
          If value IsNot Nothing Then
            ' Sync up the IProperty values using PropertyY
            Helper.AssignObjectProperties(value, Me)
            RefreshMatchStatus()
            OnPropertyChanged("PropertyY")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Set
    End Property

    Public Property Scope As PropertyScope = PropertyScope.BothDocumentAndVersionProperties

    Public Property VersionId As String = String.Empty

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpPropertyX As IProperty)
      Try
        PropertyX = lpPropertyX
        If lpPropertyX IsNot Nothing Then
          Helper.AssignObjectProperties(lpPropertyX, Me)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpPropertyX As IProperty, lpScope As PropertyScope)
      Try
        PropertyX = lpPropertyX
        If lpPropertyX IsNot Nothing Then
          Helper.AssignObjectProperties(lpPropertyX, Me)
        End If
        Scope = lpScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpPropertyX As IProperty, lpScope As PropertyScope, lpVersionId As String)
      Try
        PropertyX = lpPropertyX
        If lpPropertyX IsNot Nothing Then
          Helper.AssignObjectProperties(lpPropertyX, Me)
        End If
        Scope = lpScope
        VersionId = lpVersionId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpPropertyX As IProperty, lpPropertyY As IProperty)
      Try
        PropertyX = lpPropertyX
        PropertyY = lpPropertyY
        If lpPropertyX IsNot Nothing Then
          Helper.AssignObjectProperties(lpPropertyX, Me)
        ElseIf lpPropertyY IsNot Nothing Then
          Helper.AssignObjectProperties(lpPropertyY, Me)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IProperty Implementation"

    Public Property Cardinality As Core.Cardinality Implements Core.IProperty.Cardinality

    Public Sub Clear() Implements Core.IProperty.Clear
      Try
        Value = Nothing
        Values = New Values
        Values.SetParentProperty(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Property DefaultValue As Object Implements Core.IProperty.DefaultValue

    Public Property Description As String Implements Core.IProperty.Description

    Public Property DisplayName() As String Implements IProperty.DisplayName

    Public ReadOnly Property HasValue As Boolean Implements Core.IProperty.HasValue
      Get
        Try
          Return PropertyHasValue()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, Me.Name)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property Name As String Implements Core.IProperty.Name

    Public ReadOnly Property Persistent As Boolean Implements Core.IProperty.Persistent
      Get
        Return False
      End Get
    End Property

    Public Property SystemName As String Implements Core.IProperty.SystemName

    Public Property Type As Core.PropertyType Implements Core.IProperty.Type

    Public Property Value As Object Implements Core.IProperty.Value

    Public Property Values As Object Implements Core.IProperty.Values

#Region "Protected Methods"

    ''' <summary>
    ''' Determines whether or not the property has a value specified
    ''' </summary>
    ''' <returns>True if there is a value for the single valued case or at least one value for the multi-valued case</returns>
    ''' <remarks></remarks>
    Protected Overridable Function PropertyHasValue() As Boolean

      Try
        Select Case Me.Cardinality
          Case Core.Cardinality.ecmSingleValued
            If Value Is Nothing Then
              Return False
            End If
            If Value.ToString.Length = 0 Then
              Return False
            End If
          Case Core.Cardinality.ecmMultiValued
            If Values Is Nothing Then
              Return False
            End If
            If Values.Count = 0 Then
              Return False
            End If
            For Each lobjValue As Object In Values
              If TypeOf lobjValue Is Value Then
                If lobjValue.Value Is Nothing Then
                  Return False
                End If
              ElseIf TypeOf lobjValue Is String Then
                If String.IsNullOrEmpty(lobjValue) Then
                  Return False
                End If
              End If
            Next
        End Select

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, Me.Name)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#End Region

#Region "Public Methods"

    Public Overrides Function ToString() As String Implements IProperty.ToDebugString
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        Select Case MatchStatus
          Case MatchStatusEnum.DoesNotExistInEitherDocument
            lobjReturnBuilder.Append(MatchStatus.ToString)
          Case Else

            ' Add the scope and name
            lobjReturnBuilder.AppendFormat("{0} - {1}", Scope.ToString, Name)

            ' If this is a version property show the version id also
            If Scope = PropertyScope.VersionProperty Then
              lobjReturnBuilder.AppendFormat(" Version {0}", VersionId)
            End If

            If MatchStatus <> MatchStatusEnum.ExistsInBothDocuments Then
              lobjReturnBuilder.AppendFormat(": {0}", MatchStatus.ToString)
            End If

            If IsEqual Then
              lobjReturnBuilder.Append(" - <Equal>")
            Else
              lobjReturnBuilder.Append(" - <Not Equal>")
            End If

            lobjReturnBuilder.AppendFormat(" ({0})", Me.ComparisonValue)

            'If PropertyX.HasValue AndAlso PropertyY.HasValue Then
            '  lobjReturnBuilder.AppendFormat(" ({0} vs. {1})", PropertyX.Value.ToString, PropertyY.Value.ToString)
            'End If


            'If PropertyX IsNot Nothing Then
            '  lobjReturnBuilder.AppendFormat(" {0}", PropertyX.ToString)
            'End If

        End Select

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets a representation of the current value(s) of property x and y.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetComparisonValue() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        If PropertyX IsNot Nothing AndAlso PropertyY IsNot Nothing Then
          If PropertyX.HasValue AndAlso PropertyY.HasValue Then
            Select Case Cardinality
              Case Core.Cardinality.ecmSingleValued
                If PropertyX.Value.Equals(PropertyY.Value) Then
                  lobjReturnBuilder.AppendFormat(PropertyX.Value.ToString)
                Else
                  lobjReturnBuilder.AppendFormat("PropertyX: '{0}', PropertyY: '{1}'",
                                                 PropertyX.Value.ToString, PropertyY.Value.ToString)
                End If
              Case Else
                If PropertyX.Value.Equals(PropertyY.Value) Then
                  lobjReturnBuilder.AppendFormat(PropertyX.Value.ToString)
                Else
                  lobjReturnBuilder.AppendFormat("PropertyX: '{0}', PropertyY: '{1}'",
                                                 PropertyX.Value.ToString, PropertyY.Value.ToString)
                End If
            End Select
          Else
            If Cardinality = Core.Cardinality.ecmSingleValued Then
              lobjReturnBuilder.AppendFormat("No Value")
            Else
              lobjReturnBuilder.AppendFormat("No Values")
            End If
          End If
        Else
          If PropertyX IsNot Nothing AndAlso PropertyX.HasValue Then
            lobjReturnBuilder.AppendFormat("PropertyX: {0}, PropertyY: NoValue", PropertyX.Value.ToString)
          End If

          If PropertyY IsNot Nothing AndAlso PropertyY.HasValue Then
            lobjReturnBuilder.AppendFormat("PropertyX: NoValue, PropertyY: {0}", PropertyY.Value.ToString)
          End If

        End If

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Compares the values of both properties
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CompareProperties() As Boolean
      Try
        Select Case MatchStatus
          Case MatchStatusEnum.ExistsInBothDocuments
            If PropertyX.Cardinality <> PropertyY.Cardinality Then
              Return False
            End If
            Select Case Cardinality
              Case Core.Cardinality.ecmSingleValued
                If PropertyX.Value.Equals(PropertyY.Value) Then
                  Return True
                Else
                  Return False
                End If
              Case Else
                'If PropertyX.Values.Equals(PropertyY.Values) Then
                If PropertyX.Values = PropertyY.Values Then
                  Return True
                Else
                  Return False
                End If
            End Select
          Case Else
            Return False
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetMatchStatus() As MatchStatusEnum
      Try
        If PropertyX IsNot Nothing AndAlso PropertyY IsNot Nothing Then
          Return MatchStatusEnum.ExistsInBothDocuments
        End If

        If PropertyX Is Nothing AndAlso PropertyY Is Nothing Then
          Return MatchStatusEnum.DoesNotExistInEitherDocument
        End If

        If PropertyX IsNot Nothing AndAlso PropertyY Is Nothing Then
          Return MatchStatusEnum.DocumentXOnly
        End If

        If PropertyX Is Nothing AndAlso PropertyY IsNot Nothing Then
          Return MatchStatusEnum.DocumentYOnly
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub RefreshMatchStatus()
      Try
        menuMatchStatus = GetMatchStatus()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub OnPropertyChanged(ByVal lpPropertyName As String)
      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(lpPropertyName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      ' As per the Microsoft guidelines this is not implemented
      Return Nothing
    End Function

    Public Sub ReadXml(reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try
        mstrXsiType = reader.GetAttribute("xsi:type")
        Dim lstrCurrentElementName As String = String.Empty

        Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name = "ComparisonProperty"
          If reader.NodeType = XmlNodeType.Element Then
            lstrCurrentElementName = reader.Name
          Else
            Select Case lstrCurrentElementName
              Case "PropertyX"
                'Beep()
              Case "PropertyY"
                'Beep()
            End Select
          End If
        Loop
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          ' Write the Property Name element
          .WriteAttributeString("Name", Me.Name)

          .WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
          .WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")

          ' Write the Property IsEqual element
          .WriteAttributeString("IsEqual", Me.IsEqual.ToString)
          ' Write the Property ComparisonValue element
          .WriteAttributeString("ComparisonValue", Me.ComparisonValue)
          ' Write the Property MatchStatus element
          .WriteAttributeString("MatchStatus", Me.MatchStatus.ToString)

          ' Write the Property Scope element
          .WriteAttributeString("Scope", Me.Scope.ToString)

          ' Write the Property VersionId element
          .WriteAttributeString("VersionId", Me.VersionId)

          WritePropertyElement("PropertyX", PropertyX, writer)
          'If PropertyX IsNot Nothing Then
          '  .WriteStartElement("PropertyX")
          '  If TypeOf PropertyX Is ECMProperty Then
          '    .WriteRaw(DirectCast(PropertyX, ECMProperty).ToXmlString)
          '  End If
          '  .WriteEndElement()
          'Else
          '  .WriteElementString("PropertyX", String.Empty)
          'End If

          WritePropertyElement("PropertyY", PropertyY, writer)
          'If PropertyY IsNot Nothing Then
          '  If TypeOf PropertyY Is ECMProperty Then
          '    .WriteRaw(DirectCast(PropertyY, ECMProperty).ToXmlString)
          '  End If
          'End If

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub WritePropertyElement(lpPropertyName As String, lpProperty As IProperty, lpWriter As System.Xml.XmlWriter)
      Try
        With lpWriter
          If lpProperty IsNot Nothing Then
            .WriteStartElement(lpPropertyName)
            If TypeOf lpProperty Is ECMProperty Then
              .WriteRaw(DirectCast(lpProperty, ECMProperty).ToXmlString)
            End If
            .WriteEndElement()
          Else
            .WriteElementString(lpPropertyName, String.Empty)
          End If
        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace