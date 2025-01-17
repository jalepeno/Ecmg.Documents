'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  <Serializable(),
  XmlInclude(GetType(ClassificationXmlProperty))>
  Partial Public Class ClassificationProperty

    Public Shared Function CreateFromXmlelement(ByVal lobjPropertyElement As XmlElement) As ClassificationProperty
      Try

        Dim lobjClassificationProperty As ClassificationProperty
        Dim lobjTestElement As XmlElement
        Dim lobjChoiceListElement As XmlElement

        lobjClassificationProperty = ClassificationPropertyFactory.Create(
                               System.Enum.Parse(GetType(PropertyType), lobjPropertyElement.Item("Type").InnerText),
                               lobjPropertyElement.Item("Name").InnerText,
                               lobjPropertyElement.Item("SystemName").InnerText,
                               System.Enum.Parse(GetType(Cardinality), lobjPropertyElement.Item("Cardinality").InnerText))

        With lobjClassificationProperty
          '.Type = System.Enum.Parse(GetType(PropertyType), lobjPropertyElement.Item("Type").InnerText)
          '.Name = lobjPropertyElement.Item("Name").InnerText
          '.SystemName = lobjPropertyElement.Item("SystemName").InnerText
          '.Cardinality = System.Enum.Parse(GetType(Cardinality), lobjPropertyElement.Item("Cardinality").InnerText)

          '.IsInherited = lobjPropertyElement.Item("IsInherited").InnerText
          .IsInherited = ElementValueToBoolean(GetElementValue(lobjPropertyElement, "IsInherited"))
          '.IsRequired = lobjPropertyElement.Item("IsRequired").InnerText
          .IsRequired = ElementValueToBoolean(GetElementValue(lobjPropertyElement, "IsRequired"))
          '.IsSystemProperty = lobjPropertyElement.Item("IsSystemProperty").InnerText
          .IsSystemProperty = ElementValueToBoolean(GetElementValue(lobjPropertyElement, "IsSystemProperty"))
          '.Searchable = lobjPropertyElement.Item("Searchable").InnerText
          .DefaultValue = lobjPropertyElement.Item("DefaultValue").InnerText
          '.Selectable = lobjPropertyElement.Item("Selectable").InnerText
          .Selectable = ElementValueToBoolean(GetElementValue(lobjPropertyElement, "Selectable"))

          ' Check to see if the Settability element exists
          lobjTestElement = lobjPropertyElement.Item("Settability")

          If lobjTestElement IsNot Nothing Then
            .Settability = System.Enum.Parse(GetType(ClassificationProperty.SettabilityEnum), lobjTestElement.InnerText)
          End If

          ' Get the MaxLength
          If TypeOf lobjClassificationProperty Is ClassificationStringProperty Then
            ' Try to get the max length
            If lobjPropertyElement.Item("MaxLength") IsNot Nothing Then
              Dim lstrMaxLength As String = lobjPropertyElement.Item("MaxLength").InnerText
              If Not String.IsNullOrEmpty(lstrMaxLength) AndAlso IsNumeric(lstrMaxLength) Then
                CType(lobjClassificationProperty, ClassificationStringProperty).MaxLength = lobjPropertyElement.Item("MaxLength").InnerText
              End If
            End If

          End If

          ' Get the ChoiceList name
          lobjChoiceListElement = lobjPropertyElement.Item("ChoiceList")
          If lobjChoiceListElement IsNot Nothing Then
            .SetChoiceListName(lobjChoiceListElement.GetAttribute("name"))
          End If

        End With

        Return lobjClassificationProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ElementValueToBoolean(lpElementValue As String) As Boolean
      Try
        If String.IsNullOrEmpty(lpElementValue) Then
          Return False
        End If

        Select Case lpElementValue.ToLower
          Case "true"
            Return True
          Case "false"
            Return False
          Case Else
            Return False
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetElementValue(lpParentElement As XmlElement, lpSubElementName As String) As String
      Try
        If lpParentElement.Item(lpSubElementName) IsNot Nothing Then
          Return lpParentElement.Item(lpSubElementName).InnerText
        Else
          Return String.Empty
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a new type specific classification property based on the specified ecm property.
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(ByVal lpProperty As ECMProperty) As ClassificationProperty
      Try

        Dim lobjNewProperty As ClassificationProperty

        Select Case lpProperty.Type
          Case PropertyType.ecmString
            lobjNewProperty = New ClassificationStringProperty

          Case PropertyType.ecmBoolean
            lobjNewProperty = New ClassificationBooleanProperty
          Case PropertyType.ecmDate
            lobjNewProperty = New ClassificationDateTimeProperty
          Case PropertyType.ecmDouble
            lobjNewProperty = New ClassificationDoubleProperty
          Case PropertyType.ecmLong
            lobjNewProperty = New ClassificationLongProperty
          Case PropertyType.ecmObject
            lobjNewProperty = New ClassificationObjectProperty
          Case PropertyType.ecmGuid
            lobjNewProperty = New ClassificationGuidProperty
          Case PropertyType.ecmBinary
            lobjNewProperty = New ClassificationBinaryProperty
          Case PropertyType.ecmHtml
            lobjNewProperty = New ClassificationHtmlProperty
          Case PropertyType.ecmUri
            lobjNewProperty = New ClassificationUriProperty
          Case PropertyType.ecmXml
            lobjNewProperty = New ClassificationXmlProperty
          Case PropertyType.ecmUndefined
            lobjNewProperty = New ClassificationStringProperty
          Case Else
            lobjNewProperty = New ClassificationStringProperty
        End Select

        Helper.AssignObjectProperties(lpProperty, lobjNewProperty)
        SetDefaultValues(lobjNewProperty)

        Return lobjNewProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "ITableable Implementation"

    'Protected Overrides Function CreateDataTable() As DataTable
    '  Try
    '    ' Create New DataTable
    '    Dim propertyDataTable As New DataTable

    '    propertyDataTable.TableName = String.Format("tblProperty_{0}", Me.PackedName)

    '    ' Create columns        
    '    With propertyDataTable.Columns
    '      .Add("Name", System.Type.GetType("System.String"))
    '      .Add("SystemName", System.Type.GetType("System.String"))
    '      .Add("Cardinality", System.Type.GetType("System.String"))
    '      .Add("Type", System.Type.GetType("System.String"))
    '      .Add("Required", System.Type.GetType("System.Boolean"))
    '      .Add("Hidden", System.Type.GetType("System.Boolean"))
    '      .Add("System", System.Type.GetType("System.Boolean"))
    '      .Add("Settability", System.Type.GetType("System.String"))
    '      .Add("ChoiceList", System.Type.GetType("System.String"))
    '      .Add("DefaultValue", System.Type.GetType("System.String"))
    '      .Add("MaxLength", System.Type.GetType("System.Integer"))
    '      .Add("MinValue", System.Type.GetType("System.Integer"))
    '      .Add("MaxValue", System.Type.GetType("System.Integer"))
    '    End With

    '    Dim lstrCardinality As String = String.Empty
    '    Dim lstrSettability As String = String.Empty
    '    Dim lstrType As String = String.Empty
    '    Dim lstrChoiceListName As String = String.Empty
    '    Dim lstrDefaultValue As String = String.Empty
    '    Dim lintMaxLength As Nullable(Of Integer)
    '    Dim lintMinValue As Nullable = Nothing
    '    Dim lintMaxValue As Nullable = Nothing

    '    If Me.Cardinality = Core.Cardinality.ecmMultiValued Then
    '      lstrCardinality = "Multi Valued"
    '    Else
    '      lstrCardinality = "Single Valued"
    '    End If

    '    Select Case Settability
    '      Case SettabilityEnum.READ_ONLY
    '        lstrSettability = "Read Only"
    '      Case SettabilityEnum.READ_WRITE
    '        lstrSettability = "Read Write"
    '      Case SettabilityEnum.SETTABLE_ONLY_BEFORE_CHECKIN
    '        lstrSettability = "Only Before Checkin"
    '      Case SettabilityEnum.SETTABLE_ONLY_ON_CREATE
    '        lstrSettability = "Only on Create"
    '    End Select

    '    lstrType = Me.Type.ToString.TrimStart("ecm")

    '    If Me.ChoiceList IsNot Nothing Then
    '      lstrChoiceListName = Me.ChoiceList.Name
    '    End If

    '    If Me.DefaultValue IsNot Nothing Then
    '      lstrDefaultValue = Me.DefaultValue.ToString
    '    End If

    '    If TypeOf Me Is ClassificationStringProperty Then
    '      lintMaxLength = CType(Me, ClassificationStringProperty).MaxLength
    '    End If

    '    If (TypeOf Me Is ClassificationLongProperty) OrElse _
    '      (TypeOf Me Is ClassificationDoubleProperty) OrElse _
    '      (TypeOf Me Is ClassificationDateTimeProperty) Then
    '      lintMinValue = CType(Me, Object).MinValue
    '      lintMaxValue = CType(Me, Object).MaxValue
    '    End If

    '    propertyDataTable.Rows.Add(Me.Name, Me.SystemName, lstrCardinality, _
    '                               lstrType, Me.IsRequired, Me.IsHidden, _
    '                               Me.IsSystemProperty, lstrSettability, _
    '                               lstrChoiceListName, lstrDefaultValue, lintMaxLength, _
    '                               lintMinValue, lintMaxValue)

    '    Return propertyDataTable

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Overrides Function ToDataTable() As DataTable
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Overrides Sub ReadXml(ByVal reader As System.Xml.XmlReader)
      Try

        Dim lstrCurrentElementName As String = String.Empty

        ID = reader.GetAttribute("ID")
        Description = reader.GetAttribute("Description")

        Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name.EndsWith("Property")
          If reader.NodeType = XmlNodeType.Element Then
            lstrCurrentElementName = reader.Name
          ElseIf reader.NodeType = XmlNodeType.Text Then

            Select Case lstrCurrentElementName
              Case "Type"
                Type = [Enum].Parse(Type.GetType, reader.Value, True)
              Case "Cardinality"
                Cardinality = [Enum].Parse(Cardinality.GetType, reader.Value, True)
              Case "Name"
                Name = reader.Value
              Case "SystemName"
                SystemName = reader.Value
              Case "IsInherited"
                IsInherited = reader.Value
              Case "IsRequired"
                IsRequired = reader.Value
              Case "IsHidden"
                IsHidden = reader.Value
              Case "IsSystemProperty"
                IsSystemProperty = reader.Value
              Case "Searchable"
                Searchable = reader.Value
              Case "Settability"
                Settability = [Enum].Parse(Settability.GetType, reader.Value, True)
              Case "Selectable"
                Selectable = reader.Value
              Case "DefaultValue"
                If TypeOf Me Is ClassificationUriProperty Then
                  DirectCast(Me, ClassificationUriProperty).DefaultValueString = reader.Value
                ElseIf TypeOf Me Is ClassificationXmlProperty Then
                  DirectCast(Me, ClassificationXmlProperty).DefaultValueString = reader.Value
                Else
                  DefaultValue = reader.Value
                End If
              Case "MinValue"
                If TypeOf Me Is ClassificationDateTimeProperty Then
                  DirectCast(Me, ClassificationDateTimeProperty).MinValue = reader.Value
                ElseIf TypeOf Me Is ClassificationDoubleProperty Then
                  DirectCast(Me, ClassificationDoubleProperty).MinValue = reader.Value
                ElseIf TypeOf Me Is ClassificationLongProperty Then
                  DirectCast(Me, ClassificationLongProperty).MinValue = reader.Value
                End If
              Case "MaxValue"
                If TypeOf Me Is ClassificationDateTimeProperty Then
                  DirectCast(Me, ClassificationDateTimeProperty).MaxValue = reader.Value
                ElseIf TypeOf Me Is ClassificationDoubleProperty Then
                  DirectCast(Me, ClassificationDoubleProperty).MaxValue = reader.Value
                ElseIf TypeOf Me Is ClassificationLongProperty Then
                  DirectCast(Me, ClassificationLongProperty).MaxValue = reader.Value
                End If
              Case "MaxLength"
                If TypeOf Me Is ClassificationStringProperty Then
                  DirectCast(Me, ClassificationStringProperty).MaxLength = reader.Value
                End If
            End Select
          End If
          reader.Read()
          Do Until reader.NodeType <> XmlNodeType.Whitespace
            reader.Read()
          Loop
        Loop

        If TypeOf Me Is ClassificationDateTimeProperty Then

        End If
        reader.ReadEndElement()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Sub WriteXml(ByVal writer As System.Xml.XmlWriter)
      Try

        WriteStandardElements(writer)

        With writer
          If TypeOf Me Is ClassificationUriProperty Then
            ' Write the DefaultValue element
            .WriteElementString("DefaultValue", DirectCast(Me, ClassificationUriProperty).DefaultValueString)
          ElseIf TypeOf Me Is ClassificationXmlProperty Then
            ' Write the DefaultValue element
            .WriteElementString("DefaultValue", DirectCast(Me, ClassificationXmlProperty).DefaultValueString)
          Else
            ' Write the DefaultValue element
            '.WriteElementString("DefaultValue", DefaultValue)
            If DefaultValue IsNot Nothing Then
              .WriteElementString("DefaultValue", DefaultValue.ToString)
            Else
              .WriteElementString("DefaultValue", String.Empty)
            End If
          End If


          If TypeOf Me Is ClassificationDateTimeProperty Then
            If DirectCast(Me, ClassificationDateTimeProperty).MinValue IsNot Nothing Then
              ' Write the MinValue element
              .WriteElementString("MinValue", DirectCast(Me, ClassificationDateTimeProperty).MinValue)
            Else
              ' Write the MinValue element
              .WriteElementString("MinValue", String.Empty)
            End If
            If DirectCast(Me, ClassificationDateTimeProperty).MaxValue IsNot Nothing Then
              ' Write the MinValue element
              .WriteElementString("MaxValue", DirectCast(Me, ClassificationDateTimeProperty).MaxValue)
            Else
              ' Write the MinValue element
              .WriteElementString("MaxValue", String.Empty)
            End If
          ElseIf TypeOf Me Is ClassificationDoubleProperty Then
            If DirectCast(Me, ClassificationDoubleProperty).MinValue IsNot Nothing Then
              ' Write the MinValue element
              .WriteElementString("MinValue", DirectCast(Me, ClassificationDoubleProperty).MinValue)
            Else
              ' Write the MinValue element
              .WriteElementString("MinValue", String.Empty)
            End If
            If DirectCast(Me, ClassificationDoubleProperty).MaxValue IsNot Nothing Then
              ' Write the MaxValue element
              .WriteElementString("MaxValue", DirectCast(Me, ClassificationDoubleProperty).MaxValue)
            Else
              ' Write the MaxValue element
              .WriteElementString("MaxValue", String.Empty)
            End If
          ElseIf TypeOf Me Is ClassificationLongProperty Then
            If DirectCast(Me, ClassificationLongProperty).MinValue IsNot Nothing Then
              ' Write the MinValue element
              .WriteElementString("MinValue", DirectCast(Me, ClassificationLongProperty).MinValue)
            Else
              ' Write the MinValue element
              .WriteElementString("MinValue", String.Empty)
            End If
            If DirectCast(Me, ClassificationLongProperty).MaxValue IsNot Nothing Then
              ' Write the MaxValue element
              .WriteElementString("MaxValue", DirectCast(Me, ClassificationLongProperty).MaxValue)
            Else
              ' Write the MaxValue element
              .WriteElementString("MaxValue", String.Empty)
            End If
          End If

          If TypeOf Me Is ClassificationStringProperty Then
            If DirectCast(Me, ClassificationStringProperty).MaxLength IsNot Nothing Then
              ' Write the MaxLength element
              .WriteElementString("MaxLength", DirectCast(Me, ClassificationStringProperty).MaxLength)
            Else
              ' Write the MaxLength element
              .WriteElementString("MaxLength", String.Empty)
            End If

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