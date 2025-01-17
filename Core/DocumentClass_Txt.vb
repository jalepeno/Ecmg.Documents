'---------------------------------------------------------------------------------
' <copyright company="Conteage">
'     Copyright (c) Conteage Corp All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
'     August 14, 2024 5:34AM
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports Documents.Core.ChoiceLists
Imports Documents.Utilities

#End Region

Namespace Core

  Partial Public Class DocumentClass

    Private Const STANDARD_TXT_HEADER As String = "Name,SystemName,Cardinality,Type,Required,Hidden,System,Settability,ChoiceList,DefaultValue,MaxLength,MinValue,MaxValue"

    ''' <summary>
    ''' Creates a DocumentClass object from a text file
    ''' </summary>
    ''' <param name="lpFilePath"></param>
    ''' <returns></returns>
    Public Shared Function FromTxtFile(lpFilePath As String) As DocumentClass
      Try
        Dim lobjDocumentClass As New DocumentClass

        Dim lstrDocumentClassLines As String() = File.ReadAllLines(lpFilePath)
        Dim lintLineCounter As Integer = 0

        ' Read the DocumentClass Name
        For Each lstrLine As String In lstrDocumentClassLines
          lintLineCounter += 1
          Select Case lintLineCounter
            Case 1
              lobjDocumentClass.Name = lstrLine
              lobjDocumentClass.Label = lstrLine
            Case 2
              ' Skip this line, it is the column headers
              Continue For
            Case Else
              lobjDocumentClass.Properties.Add(ReadClassificationPropertyLine(lstrLine))
          End Select
        Next


        Return lobjDocumentClass

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ReadClassificationPropertyLine(lpLine As String) As ClassificationProperty
      Try
        Dim lstrColumns As String() = lpLine.Split(",")
        Dim lintColumnCounter As Integer = 0

        Dim lstrName As String = String.Empty
        Dim lstrSystemName As String = String.Empty
        Dim lstrCardinality As String = String.Empty
        Dim lenuCardinality As Cardinality
        Dim lstrType As String = String.Empty
        Dim lenumType As PropertyType
        Dim lstrRequired As String = String.Empty
        Dim lstrHidden As String = String.Empty
        Dim lstrSystem As String = String.Empty
        Dim lstrSettability As String = String.Empty
        Dim lstrChoiceList As String = String.Empty
        Dim lstrDefaultValue As String = String.Empty
        Dim lstrMaxLength As String = String.Empty
        Dim lstrMinValue As String = String.Empty
        Dim lstrMaxValue As String = String.Empty


        For Each lstrColumn As String In lstrColumns
          lintColumnCounter += 1
          Select Case lintColumnCounter
            Case 1
              lstrName = lstrColumn
            Case 2
              lstrSystemName = lstrColumn
            Case 3
              lstrCardinality = lstrColumn
            Case 4
              lstrType = lstrColumn
            Case 5
              lstrRequired = lstrColumn
            Case 6
              lstrHidden = lstrColumn
            Case 7
              lstrSystem = lstrColumn
            Case 8
              lstrSettability = lstrColumn
            Case 9
              lstrChoiceList = lstrColumn
            Case 10
              lstrDefaultValue = lstrColumn
            Case 11
              If IsNumeric(lstrColumn) Then
                lstrMaxLength = lstrColumn
              Else
                lstrMaxLength = String.Empty
              End If
            Case 12
              If IsNumeric(lstrColumn) Then
                lstrMinValue = lstrColumn
              Else
                lstrMinValue = String.Empty
              End If
            Case 13
              If IsNumeric(lstrColumn) Then
                lstrMaxValue = lstrColumn
              Else
                lstrMaxValue = String.Empty
              End If
          End Select
        Next

        lenuCardinality = CardinalityFromString(lstrCardinality)
        lenumType = PropertyTypeFromString(lstrType)

        Dim lobjProperty As ClassificationProperty = ClassificationPropertyFactory.Create(lenumType, lstrName, lstrSystemName, lenuCardinality)

        With lobjProperty

          ' Set IsRequired
          .IsRequired = Boolean.Parse(lstrRequired)

          ' Set IsHidden
          .IsHidden = Boolean.Parse(lstrHidden)

          ' Set Is SystemProperty
          .IsSystemProperty = Boolean.Parse(lstrSystem)

          ' Set Settability
          .Settability = SettabilityFromString(lstrSettability)

          ' Set ChoiceList (at least the name)
          If Not String.IsNullOrEmpty(lstrChoiceList) Then
            .ChoiceList = New ChoiceList()
            .ChoiceList.Name = lstrChoiceList
          End If

          ' Set Default Value
          If Not String.IsNullOrEmpty(lstrDefaultValue) Then
            .DefaultValue = lstrDefaultValue
          End If

          ' Set Max Length
          If Not String.IsNullOrEmpty(lstrMaxLength) Then
            If lenumType = PropertyType.ecmString Then
              DirectCast(lobjProperty, ClassificationStringProperty).MaxLength = Integer.Parse(lstrMaxLength)
            Else
              ApplicationLogging.WriteLogEntry(String.Format("MaxLength is not a valid property type {0} in property {1}", lenumType.ToString(), lstrName))
            End If
          End If

          ' Set MinValue
          If Not String.IsNullOrEmpty(lstrMinValue) Then
            Select Case lenumType
              Case PropertyType.ecmLong
                DirectCast(lobjProperty, ClassificationLongProperty).MinValue = Integer.Parse(lstrMinValue)
              Case PropertyType.ecmDouble
                DirectCast(lobjProperty, ClassificationDoubleProperty).MinValue = Integer.Parse(lstrMinValue)
              Case Else
                ApplicationLogging.WriteLogEntry(String.Format("MinValue is not a valid property type {0} in property {1}", lenumType.ToString(), lstrName))
            End Select
          End If

          ' Set MaxValue
          If Not String.IsNullOrEmpty(lstrMaxValue) Then
            Select Case lenumType
              Case PropertyType.ecmLong
                DirectCast(lobjProperty, ClassificationLongProperty).MaxValue = Integer.Parse(lstrMaxValue)
              Case PropertyType.ecmDouble
                DirectCast(lobjProperty, ClassificationDoubleProperty).MaxValue = Integer.Parse(lstrMaxValue)
              Case Else
                ApplicationLogging.WriteLogEntry(String.Format("MinValue is not a valid property type {0} in property {1}", lenumType.ToString(), lstrName))
            End Select
          End If
        End With

        Return lobjProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function PropertyTypeFromString(lpType As String) As PropertyType
      Try
        Select Case lpType
          Case "string"
            Return PropertyType.ecmString
          Case "bool"
            Return PropertyType.ecmBoolean
          Case "date"
            Return PropertyType.ecmDate
          Case "double"
            Return PropertyType.ecmDouble
          Case "long"
            Return PropertyType.ecmLong
          Case "binary"
            Return PropertyType.ecmBinary
          Case "enum"
            Return PropertyType.ecmEnum
          Case "guid"
            Return PropertyType.ecmGuid
          Case "html"
            Return PropertyType.ecmHtml
          Case "object"
            Return PropertyType.ecmObject
          Case "uri"
            Return PropertyType.ecmUri
          Case "valuemap"
            Return PropertyType.ecmValueMap
          Case "xml"
            Return PropertyType.ecmXml
          Case "unknown"
            Return PropertyType.ecmUndefined
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function StringFromPropertyType(lpType As PropertyType) As String
      Try
        Select Case lpType
          Case PropertyType.ecmString
            Return "string"

          Case PropertyType.ecmBoolean
            Return "bool"

          Case PropertyType.ecmDate
            Return "date"

          Case PropertyType.ecmDouble
            Return "double"

          Case PropertyType.ecmLong
            Return "long"

          Case PropertyType.ecmBinary
            Return "binary"

          Case PropertyType.ecmEnum
            Return "enum"

          Case PropertyType.ecmGuid
            Return "guid"

          Case PropertyType.ecmHtml
            Return "html"

          Case PropertyType.ecmObject
            Return "object"

          Case PropertyType.ecmUri
            Return "uri"

          Case PropertyType.ecmValueMap
            Return "valuemap"

          Case PropertyType.ecmXml
            Return "xml"

          Case Else
            Return "unknown"

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function SettabilityFromString(lpSettability As String) As ClassificationProperty.SettabilityEnum
      Try
        Select Case lpSettability
          Case "readonly"
            Return ClassificationProperty.SettabilityEnum.READ_ONLY

          Case "readwrite"
            Return ClassificationProperty.SettabilityEnum.READ_WRITE

          Case "settableonlybeforecheckin"
            Return ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_BEFORE_CHECKIN

          Case "settableonlyoncreate"
            Return ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE

        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function StringFromSettability(lpSettability As ClassificationProperty.SettabilityEnum) As String
      Try

        Select Case lpSettability
          Case ClassificationProperty.SettabilityEnum.READ_ONLY
            Return "readonly"

          Case ClassificationProperty.SettabilityEnum.READ_WRITE
            Return "readwrite"

          Case ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_BEFORE_CHECKIN
            Return "settableonlybeforecheckin"

          Case ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE
            Return "settableonlyoncreate"

        End Select

        Return "unknown"

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function CardinalityFromString(lpCardinality As String) As Cardinality
      Try
        Select Case lpCardinality
          Case "sv"
            Return Cardinality.ecmSingleValued

          Case "mv"
            Return Cardinality.ecmMultiValued

          Case Else
            Return Cardinality.ecmSingleValued
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function StringFromCardinality(lpCardinality As Cardinality) As String
      Try
        Select Case lpCardinality
          Case Cardinality.ecmSingleValued
            Return "sv"

          Case Cardinality.ecmMultiValued
            Return "mv"

        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves the DocumentClass object to a text file
    ''' </summary>
    ''' <param name="lpFilePath"></param>
    ''' <returns></returns>
    Public Sub ToTxtFile(lpFilePath As String)
      Try
        Dim lobjClassStringBuilder As New StringBuilder

        ' Add the Class Name
        lobjClassStringBuilder.AppendLine(Name)

        'Add the header line
        lobjClassStringBuilder.AppendLine(STANDARD_TXT_HEADER)

        For Each lobjProperty As ClassificationProperty In Properties
          Dim lobjPropertyStringBuilder As New StringBuilder
          ' Append the Name and SystemName
          lobjPropertyStringBuilder.AppendFormat("{0},{1},", lobjProperty.Name, lobjProperty.SystemName)

          ' Append the Cardinality
          lobjPropertyStringBuilder.AppendFormat("{0},", StringFromCardinality(lobjProperty.Cardinality))

          ' Append the Type
          lobjPropertyStringBuilder.AppendFormat("{0},", StringFromPropertyType(lobjProperty.Type))

          ' Append Required
          lobjPropertyStringBuilder.AppendFormat("{0},", lobjProperty.IsRequired.ToString().ToLower())

          ' Append Hidden
          lobjPropertyStringBuilder.AppendFormat("{0},", lobjProperty.IsHidden.ToString().ToLower())

          ' Append System
          lobjPropertyStringBuilder.AppendFormat("{0},", lobjProperty.IsSystemProperty.ToString().ToLower())

          ' Append Settability
          lobjPropertyStringBuilder.AppendFormat("{0}", StringFromSettability(lobjProperty.Settability))

          ' Add the choice list, if exists
          If Not String.IsNullOrEmpty(lobjProperty.ChoiceListName) Then
            lobjPropertyStringBuilder.AppendFormat(",{0}", lobjProperty.ChoiceListName)
          Else
            lobjPropertyStringBuilder.Append(",")
          End If

          ' Add the default value, if exists
          If lobjProperty.DefaultValue IsNot Nothing Then
            lobjPropertyStringBuilder.AppendFormat(",{0}", lobjProperty.DefaultValue.ToString())
          Else
            lobjPropertyStringBuilder.Append(",")
          End If

          ' Add the max length, if exists
          If lobjProperty.Type = PropertyType.ecmString Then
            If DirectCast(lobjProperty, ClassificationStringProperty).MaxLength.HasValue Then
              lobjPropertyStringBuilder.AppendFormat(",{0}", DirectCast(lobjProperty, ClassificationStringProperty).MaxLength)
            Else
              lobjPropertyStringBuilder.Append(",")
            End If
          Else
            lobjPropertyStringBuilder.Append(",")
          End If

          ' Add the MinValue and MaxValue, if exists
          Select Case lobjProperty.Type
            Case PropertyType.ecmLong
              ' Add the MinValue, if exists
              If DirectCast(lobjProperty, ClassificationLongProperty).MinValue.HasValue Then
                lobjPropertyStringBuilder.AppendFormat(",{0}", DirectCast(lobjProperty, ClassificationLongProperty).MinValue)
              Else
                lobjPropertyStringBuilder.Append(",")
              End If

              ' Add the MaxValue, if exists
              If DirectCast(lobjProperty, ClassificationLongProperty).MaxValue.HasValue Then
                lobjPropertyStringBuilder.AppendFormat(",{0}", DirectCast(lobjProperty, ClassificationLongProperty).MaxValue)
              Else
                lobjPropertyStringBuilder.Append(",")
              End If

            Case PropertyType.ecmDouble
              ' Add the MinValue, if exists
              If DirectCast(lobjProperty, ClassificationDoubleProperty).MinValue.HasValue Then
                lobjPropertyStringBuilder.AppendFormat(",{0}", DirectCast(lobjProperty, ClassificationDoubleProperty).MinValue)
              Else
                lobjPropertyStringBuilder.Append(",")
              End If

              ' Add the MaxValue, if exists
              If DirectCast(lobjProperty, ClassificationDoubleProperty).MaxValue.HasValue Then
                lobjPropertyStringBuilder.AppendFormat(",{0}", DirectCast(lobjProperty, ClassificationDoubleProperty).MaxValue)
              Else
                lobjPropertyStringBuilder.Append(",")
              End If

          End Select

          ' Trim off any trailing commas
          Dim lstrProperty As String = lobjPropertyStringBuilder.ToString().TrimEnd(",")

          ' Add the property to the class string output
          lobjClassStringBuilder.AppendLine(lstrProperty)

        Next

        If File.Exists(lpFilePath) Then
          File.Delete(lpFilePath)
        End If

        ' Write the file
        File.AppendAllText(lpFilePath, lobjClassStringBuilder.ToString())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' Sample Text File Output
    ' Note the header on line 2

    'File System Document
    'Name,SystemName,Cardinality,Type,Required,Hidden,System,Settability,ChoiceList,DefaultValue,MaxLength,MinValue,MaxValue
    'Archive,Archive,sv,bool,false,false,true,readwrite
    'Compressed,Compressed,sv,bool,false,false,true,readwrite
    'DateCreated,DateCreated,sv,date,false,false,true,readwrite
    'DateLastAccessed,DateLastAccessed,sv,date,false,false,true,readwrite
    'DateLastModified,DateLastModified,sv,date,false,false,true,readwrite
    'Device,Device,sv,bool,false,false,true,readwrite
    'Encrypted,Encrypted,sv,bool,false,false,true,readwrite
    'FileName,FileName,sv,string,false,false,true,readwrite,,,260
    'Folder Path, Folder Path,mv,String,False,False,True,readwrite
    'Hidden,Hidden,sv,bool,false,false,true,readwrite
    'Normal,Normal,sv,bool,false,false,true,readwrite
    'NotContentIndexed,NotContentIndexed,sv,bool,false,false,true,readwrite
    'Offline,Offline,sv,bool,false,false,true,readwrite
    'Path,Path,sv,string,false,false,true,readwrite,,,248
    'ReadOnly,ReadOnly,sv,bool,false,false,true,readwrite
    'ReparsePoint,ReparsePoint,sv,bool,false,false,true,readwrite
    'SparseFile,SparseFile,sv,bool,false,false,true,readwrite
    'System,System,sv,bool,false,false,true,readwrite
    'Temporary,Temporary,sv,bool,false,false,true,readwrite

  End Class

End Namespace

