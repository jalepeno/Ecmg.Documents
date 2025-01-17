'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core.ChoiceLists

  Partial Public Class ChoiceList
    Implements ISerialize
    Implements ITableable

#Region "Constructors"

    ''' <summary>
    ''' Constructs a new choice list object using the specified choice list file.
    ''' </summary>
    ''' <param name="lpXMLFilePath">The fully qualified path to the choice list file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpXMLFilePath As String)
      Try

        Dim lobjChoiceList As ChoiceList = Deserialize(lpXMLFilePath)
        Helper.AssignObjectProperties(lobjChoiceList, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Helper.DumpException(ex)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

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
        Return CHOICE_LIST_FILE_EXTENSION
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
    ''' Saves a representation of the Document object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize

      ' Set the extension if necessary
      ' We want the default file extension on a ChoiceList to be .cvl for
      ' "Controlled Vocabulary List"

      If lpFileExtension.Length = 0 Then
        ' No override was provided
        If lpFilePath.EndsWith(CHOICE_LIST_FILE_EXTENSION) = False Then
          lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & CHOICE_LIST_FILE_EXTENSION
        End If

      End If

      Serializer.Serialize.XmlFile(Me, lpFilePath)
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      If lpWriteProcessingInstruction = True Then
        Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      End If
    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Return Serializer.Serialize.Xml(Me)
    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Return Serializer.Serialize.XmlString(Me)
    End Function

#End Region

#Region "Public Methods"

    Public Sub LoadFromFile(ByVal lpXMLFilePath As String)
      Try
        Dim lobjXmlDocument As New Xml.XmlDocument
        lobjXmlDocument.Load(lpXMLFilePath)

        LoadFromXmlDocument(lobjXmlDocument)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
      End Try
    End Sub

    Public Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
      Try
        Dim lobjChoiceList As ChoiceList = DeSerialize(lpXML)
        With Me
          .Id = lobjChoiceList.Id
          .Name = lobjChoiceList.Name
          .DescriptiveText = lobjChoiceList.DescriptiveText
          .DisplayName = lobjChoiceList.DisplayName
          .ChoiceValues = lobjChoiceList.ChoiceValues
        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
      End Try
    End Sub

    Public Shared Function CreateFromTextFile(lpSourceFilePath As String) As ChoiceList
      Try

        Helper.VerifyFilePath(lpSourceFilePath, True)
        Dim lstrName As String = IO.Path.GetFileNameWithoutExtension(lpSourceFilePath)
        Dim lobjChoiceValues As IList(Of String) = Helper.ReadTextList(lpSourceFilePath)

        Return CreateFromList(lobjChoiceValues, lstrName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateFromExcelFile(lpSourceFilePath As String) As ChoiceList
      Try

        Helper.VerifyFilePath(lpSourceFilePath, True)
        Dim lstrName As String = IO.Path.GetFileNameWithoutExtension(lpSourceFilePath)
        Dim lobjChoiceValues As IList(Of String) = Helper.ReadExcelList(lpSourceFilePath)

        Return CreateFromList(lobjChoiceValues, lstrName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Public Shared Function CreateFromList(lplist As IList(Of String), lpName As String) As ChoiceList
      Try

        Dim lobjChoiceList As New ChoiceList
        lobjChoiceList.Name = lpName
        For Each lstrValue As String In lplist
          lobjChoiceList.ChoiceValues.Add(New ChoiceValue(lstrValue))
        Next

        Return lobjChoiceList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As System.Data.DataTable Implements ITableable.ToDataTable
      Try

        ' Create New DataTable
        Dim lobjDataTable As New DataTable(String.Format("tbl{0}_{1}", Me.GetType.Name, Me.Name))

        ' Create columns        
        With lobjDataTable.Columns
          .Add("Name", System.Type.GetType("System.String"))
          .Add("DisplayName", System.Type.GetType("System.String"))
          .Add("DescriptiveText", System.Type.GetType("System.String"))
        End With

        For Each lobjChoiceItem As ChoiceItem In Me.ChoiceValues
          lobjDataTable.Rows.Add(lobjChoiceItem.Name, lobjChoiceItem.DisplayName, lobjChoiceItem.DescriptiveText)
        Next

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace