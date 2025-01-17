'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Data
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core
  ''' <summary>Defines a document or version property.</summary>
  <TypeConverter(GetType(ExpandableObjectConverter)),
  JsonConverter(GetType(EcmPropertyConverter)),
  Serializable(),
  XmlInclude(GetType(SingletonProperty)),
  XmlInclude(GetType(MultiValueProperty))>
  Partial Public Class ECMProperty
    Implements IComparable
    Implements ISerialize
    Implements ICloneable
    Implements IEquatable(Of ECMProperty)
    Implements ITableable
    'Implements ILoggable

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const ECM_PROPERTY_FILE_EXTENSION As String = "epf"

#End Region

#Region "Public Methods"

    Public Function AddValue(ByVal lpValue As Object, ByVal lpAllowDuplicates As Boolean) As Boolean
      Try
        Return AddValue(lpValue, Me, lpAllowDuplicates)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function AddValue(ByVal lpValue As Object, ByVal lpDestinationProperty As ECMProperty, ByVal lpAllowDuplicates As Boolean) As Boolean
      Try

        Dim lstrErrorMessage As String = String.Empty

        ' Make sure the destination property is multi-valued
        If lpDestinationProperty.Cardinality <> Cardinality.ecmMultiValued Then
          lstrErrorMessage = String.Format("Unable to AddValue, the destination property '{0}' is a singlevalued property, this action is only valid for multi-valued destination properties.", lpDestinationProperty.Name)
          Throw New Exceptions.InvalidCardinalityException(lstrErrorMessage, Cardinality.ecmMultiValued, lpDestinationProperty)
          ApplicationLogging.WriteLogEntry(lstrErrorMessage, TraceEventType.Error)
        End If

        Select Case lpDestinationProperty.Type
          Case PropertyType.ecmString
            lpDestinationProperty.Values.AddString(lpValue, lpAllowDuplicates)
          Case PropertyType.ecmBoolean
            lpDestinationProperty.Values.AddBoolean(lpValue, lpAllowDuplicates)
          Case PropertyType.ecmDate
            lpDestinationProperty.Values.AddDate(lpValue, lpAllowDuplicates)
          Case PropertyType.ecmDouble
            lpDestinationProperty.Values.AddDouble(lpValue, lpAllowDuplicates)
          Case PropertyType.ecmGuid
            lpDestinationProperty.Values.AddString(lpValue, lpAllowDuplicates)
          Case PropertyType.ecmLong
            lpDestinationProperty.Values.AddLong(lpValue, lpAllowDuplicates)
          Case Else
            ApplicationLogging.WriteLogEntry(
              String.Format("Could not add value for property '{0}', data type {1} not supported for operation.",
                            lpDestinationProperty.Name, lpDestinationProperty.Type.ToString.Substring(3)),
                          Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 62384)
            Return False
        End Select
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ChangePropertyValue(ByVal lpNewValue As Object) As Boolean

      Try
        Dim lstrDebugPattern As String
        ' Is the property SingleValued or MultiValued?
        If Cardinality = Cardinality.ecmSingleValued Then
          ' Set the SingleValued property value
          lstrDebugPattern = String.Format("ECMProperty({0}).ChangePropertyValue {1}: ", Me.Name, Me.GetType.Name)
          lstrDebugPattern = lstrDebugPattern & "OriginalValue = {0}, NewValue = {1}"

          ' In some cases we may need to cast the value
          Select Case Type
            Case PropertyType.ecmString
              If lpNewValue Is Nothing Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, String.Empty)
                Value = String.Empty
              ElseIf TypeOf (lpNewValue) Is DBNull Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, String.Empty)
                Value = String.Empty
              Else
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = lpNewValue
              End If
            Case PropertyType.ecmDate
              If TypeOf (lpNewValue) Is String Then
                ' Let's try to make it a date
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = DateTime.Parse(lpNewValue)
              ElseIf TypeOf (lpNewValue) Is DBNull Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, DateTime.MinValue)
                Value = DateTime.MinValue
              Else
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = lpNewValue
              End If
            Case PropertyType.ecmBoolean
              If lpNewValue Is Nothing Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, False)
                Value = False
              ElseIf TypeOf (lpNewValue) Is DBNull Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, False)
                Value = False
              Else
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = lpNewValue
              End If
            Case PropertyType.ecmLong, PropertyType.ecmDouble
              If lpNewValue Is Nothing Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, 0)
                Value = 0
              ElseIf TypeOf (lpNewValue) Is DBNull Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, 0)
                Value = 0
              ElseIf IsNumeric(lpNewValue) = False Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, 0)
                Value = 0
              Else
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = lpNewValue
              End If
            Case Else
              If TypeOf (lpNewValue) Is DBNull Then
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, String.Empty)
                Value = Nothing
              Else
                'LogSession.LogDebug(lstrDebugPattern, Me.ValueString, lpNewValue)
                Value = lpNewValue
              End If
          End Select

        ElseIf Cardinality = Cardinality.ecmMultiValued Then
          lstrDebugPattern = "ECMProperty({0}).ChangePropertyValue MultiValued: OriginalValue = {1}, NewValue = {2}"
          ' Set the MultiValued property value
          'Values = lpNewValue
          'lpNewValue is a MapList
          Dim lobjMapList As Transformations.MapList = lpNewValue
          For Each lobjValueMap As Transformations.ValueMap In lobjMapList
            For i As Integer = 0 To Values.Count - 1
              If StrComp(Values(i).ToString, lobjValueMap.Original, CompareMethod.Text) = 0 Then
                Try
                  'LogSession.LogDebug(lstrDebugPattern, Values(i), lobjValueMap.Replacement)
                Catch ex As Exception
                  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                  ' Ignore the logging error and move on.
                End Try
                Values(i) = lobjValueMap.Replacement
              End If
            Next
          Next
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False

      End Try

      Return True

    End Function

    Public Sub ClearPropertyValue()
      Try
        ' Is the property SingleValued or MultiValued?
        If Cardinality = Cardinality.ecmSingleValued Then
          Select Case Type
            Case PropertyType.ecmString
              Value = String.Empty
            Case PropertyType.ecmDate
              Value = DateTime.MinValue
            Case PropertyType.ecmBoolean
              Value = False
            Case PropertyType.ecmLong, PropertyType.ecmDouble
              Value = 0
            Case Else
              Value = Nothing
          End Select
        Else
          Values.Clear()
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "ITableable Implementation"

    Public Overridable Function ToDataTable() As DataTable Implements ITableable.ToDataTable

      Try

        ' Create New DataTable
        Dim propertyDataTable As DataTable = CreateDataTable()

        Return FillDataTable(propertyDataTable)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

    Public Function ToDelimitedString() As String
      Try
        Return ToDelimitedString(";")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToDelimitedString(lpDelimiter As String) As String
      Try
        If Cardinality = Core.Cardinality.ecmSingleValued Then
          If HasValue Then
            Return Value.ToString
          Else
            Return String.Empty
          End If
        Else
          If HasValue Then
            Dim lobjStringBuilder As New StringBuilder
            For Each lobjValue As Object In Values
              lobjStringBuilder.AppendFormat("{0};", lobjValue.ToString)
            Next
            lobjStringBuilder.Remove(lobjStringBuilder.Length - 1, 1)
            Return lobjStringBuilder.ToString
          Else
            Return String.Empty
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function CreateDataTable() As DataTable

      Try

        ' Create New DataTable
        Dim propertyDataTable As New DataTable

        propertyDataTable.TableName = String.Format("tblProperty_{0}", Me.PackedName)

        ' Create columns        
        With propertyDataTable.Columns
          .Add("Name", System.Type.GetType("System.String"))
          .Add("PackedName", System.Type.GetType("System.String"))
          .Add("ID", System.Type.GetType("System.String"))
          .Add("Type", System.Type.GetType("System.String"))
          .Add("Cardinality", System.Type.GetType("System.String"))
          .Add("DefaultValue", System.Type.GetType("System.String"))
          .Add("HasValue", System.Type.GetType("System.Boolean"))
          .Add("Value", System.Type.GetType("System.String"))
        End With

        Return propertyDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Protected Overridable Function FillDataTable(ByVal lpTable As DataTable) As DataTable

      Try

        ' Declare row        
        Dim propertyRow As DataRow

        ' create new row        
        propertyRow = lpTable.NewRow
        propertyRow("Name") = Me.Name
        propertyRow("PackedName") = Me.PackedName
        propertyRow("ID") = Me.ID
        propertyRow("Type") = Me.Type.ToString
        propertyRow("Cardinality") = Me.Cardinality.ToString
        propertyRow("DefaultValue") = Me.DefaultValue.ToString
        propertyRow("HasValue") = Me.HasValue
        propertyRow("Value") = Me.Value.ToString
        lpTable.Rows.Add(propertyRow)

        Return lpTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    '#Region "SetValue"

    '    Public Sub SetValue(ByVal lpValue As Object)
    '      Me.Value.Value = lpValue
    '    End Sub

    '#End Region

    Public Shared Operator =(ByVal v1 As ECMProperty, ByVal v2 As IProperty) As Boolean
      Try
        If ((v1 Is Nothing) AndAlso (v2 Is Nothing)) Then
          Return True
        ElseIf ((v1 Is Nothing) AndAlso (v2 IsNot Nothing)) Then
          Return False
        Else
          Return v1.Equals(v2)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator <>(ByVal v1 As ECMProperty, ByVal v2 As IProperty) As Boolean
      Try
        Return Not v1 = v2
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Overloads Function Equals(ByVal other As Object) As Boolean
      Try
        If other Is Nothing Then
          Return False
        End If
        If Not Me.GetType.IsAssignableFrom(other.GetType()) Then
          Return False
        Else
          Return Equals(CType(other, ECMProperty))
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Equals(ByVal other As ECMProperty) As Boolean Implements System.IEquatable(Of ECMProperty).Equals
      Try

        If other Is Nothing Then
          Return False
        End If

        With other

          If .Name <> Name Then
            Return False
          End If

          If .Cardinality <> Cardinality Then
            Return False
          End If

          'If .DefaultValue <> DefaultValue Then
          '  Return False
          'End If

          If .Description <> Description Then
            Return False
          End If

          If .Persistent <> Persistent Then
            Return False
          End If

          If .Type <> Type Then
            Return False
          End If

          If .Value <> Value Then
            Return False
          End If

          If .Values IsNot Nothing AndAlso Values IsNot Nothing Then
            If .Values.Count <> Values.Count Then
              Return False
            End If

            For Each lobjValue As Object In .Values
              If Values.Contains(lobjValue) = False Then
                Return False
              End If
            Next
          End If

        End With

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function StringToASCIIKeyMap(lpString As String) As List(Of String)
      Try

        Dim lobjASCIIKeyList As New List(Of String)
        For lintCounter As Integer = 0 To lpString.Length - 1
          lobjASCIIKeyList.Add(String.Format("'{0}', {1}", lpString(lintCounter), Asc(lpString(lintCounter))))
        Next

        Return lobjASCIIKeyList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"



    'Private Shared Function GetPackedName(ByVal lpInputName As String) As String
    '  Try
    '    lpInputName = lpInputName.Replace(" ", "")
    '    lpInputName = lpInputName.Replace("/", "")
    '    lpInputName = lpInputName.Replace("\", "")

    '    Return lpInputName
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Function

    ' ''' <summary>
    ' ''' Determines whether or not the property has a value specified
    ' ''' </summary>
    ' ''' <returns>True if there is a value for the single valued case or at least one value for the multi-valued case</returns>
    ' ''' <remarks></remarks>
    'Private Function PropertyHasValue() As Boolean

    '  Try
    '    Select Case Me.Cardinality
    '      Case Core.Cardinality.ecmSingleValued
    '        If Value Is Nothing Then
    '          Return False
    '        End If
    '        If Value.ToString.Length = 0 Then
    '          Return False
    '        End If
    '      Case Core.Cardinality.ecmMultiValued
    '        If Values Is Nothing Then
    '          Return False
    '        End If
    '        If Values.Count = 0 Then
    '          Return False
    '        End If
    '        For Each lstrValue As String In Values
    '          If lstrValue.Length = 0 Then
    '            Return False
    '          End If
    '        Next
    '    End Select

    '    Return True
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      If TypeOf obj Is IProperty Then
        Return PropertyComparer.CompareProperties(Me, obj)
      Else
        ' If the other object is not an IProperty object, 
        ' we will call this object greater.
        Return 1
      End If

      'If TypeOf obj Is ECMProperty Then
      '  Return Name.CompareTo(obj.Name)
      'Else
      '  Throw New ArgumentException("Object is not an ObjectProperty")
      'End If

    End Function

#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return ECM_PROPERTY_FILE_EXTENSION
      End Get
    End Property

    ''' <summary>
    ''' Instantiate from an XML file.
    ''' </summary>
    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        ' Set the extension if necessary
        ' We want the default file extension on an ecmdocument to be .cdf for
        ' "Content Definition File"

        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        Serializer.Serialize.XmlFile(Me, lpFilePath)
        'xml-stylesheet xml-stylesheet type="text/xsl" href="http://Ecmg.us/Ecmg/ecmdocument.xslt"
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
        'xml-stylesheet xml-stylesheet type="text/xsl" href="http://Ecmg.us/Ecmg/ecmdocument.xslt"
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICloneable"

    ''' <summary>
    ''' Clone a ECMProperty
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjProperty As ECMProperty = PropertyFactory.Create(Me.Type, Me.Name, Me.SystemName, Me.Cardinality)

      Try
        lobjProperty.DefaultValue = Me.DefaultValue
        lobjProperty.ID = Me.ID
        If (lobjProperty.Cardinality = Core.Cardinality.ecmSingleValued) Then
          lobjProperty.Value = Me.Value
        Else
          lobjProperty.Values = Me.Values
        End If

        lobjProperty.SetPersistence(Me.Persistent)

        Return lobjProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "ILoggable Implementation"

    'Private mobjLogSession As Gurock.SmartInspect.Session

    'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
    '  Try
    '    ApplicationLogging.FinalizeLogSession(mobjLogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.LavenderBlush)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
    '  Get
    '    Try
    '      If mobjLogSession Is Nothing Then
    '        InitializeLogSession()
    '      End If
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

#End Region

  End Class

End Namespace