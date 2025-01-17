'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

'Imports Ecmg.Cts.Core
'Imports Ecmg.Cts.SerializationUtilities
'Imports Ecmg.Cts.Utilities
Imports System.Data
Imports System.Reflection
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities


#End Region

Namespace Providers

  <XmlRoot("ProviderCatalog")>
  Public Class ProviderCatalog
    Inherits CCollection(Of ProviderInformation)
    Implements ISerialize
    'Implements ITableable

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const PROVIDER_CATALOG_FILE_EXTENSION As String = "pcf"

#End Region

#Region "Overloaded Methods"

    ''' <summary>
    ''' Checks to see if the requested provider exists in the catalog
    ''' </summary>
    ''' <param name="lpProviderName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Exists(ByVal lpProviderName As String) As Boolean
      Try
        For Each lobjProviderInfo As ProviderInformation In Me
          If lobjProviderInfo.Name = lpProviderName Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Adds a ProviderInformation object to the catalog using the provider path
    ''' </summary>
    ''' <param name="lpProviderPath"></param>
    ''' <remarks></remarks>
    Public Shadows Sub Add(ByVal lpProviderPath As String)
      Try
        MyBase.Add(New ProviderInformation(lpProviderPath))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds the specified ProviderInformation object to the catalog
    ''' </summary>
    ''' <param name="lpProviderInformation">A ProviderInformation object reference</param>
    ''' <remarks></remarks>
    Public Shadows Sub Add(ByVal lpProviderInformation As ProviderInformation)
      Try
        MyBase.Add(lpProviderInformation)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds the specified Provider to the catalog
    ''' </summary>
    ''' <param name="lpProvider">A Provider object reference</param>
    ''' <remarks>Automatically creates a ProviderInformation 
    ''' object from the specified Provider</remarks>
    Public Shadows Sub Add(ByVal lpProvider As CProvider)
      Try
        MyBase.Add(New ProviderInformation(lpProvider))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Removes the specified provider from the catalog
    ''' </summary>
    ''' <param name="lpProviderName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Remove(ByVal lpProviderName As String) As Boolean
      Try
        For Each lobjProviderInformation As ProviderInformation In Me
          If lobjProviderInformation.Name = lpProviderName Then
            Return Remove(lobjProviderInformation)
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    '''' <summary>
    '''' Gets or sets an item in the catalog
    '''' </summary>
    '''' <param name="Name">The Provider name</param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Default Shadows Property Item(ByVal Name As String) As ProviderInformation
    '  Get
    '    Try
    '      ApplicationLogging.WriteLogEntry("Enter ProviderCatalog::Get_Item(Name)", TraceEventType.Verbose)
    '      Dim lobjProviderInformation As ProviderInformation
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjProviderInformation = CType(MyBase.Item(lintCounter), ProviderInformation)
    '        If lobjProviderInformation.Name = Name Then
    '          Return lobjProviderInformation
    '        End If
    '      Next
    '      Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, "ProviderCatalog::Get_Item(Name)")
    '      Return Nothing
    '    Finally
    '      ApplicationLogging.WriteLogEntry("Exit ProviderCatalog::Get_Item(Name)", TraceEventType.Verbose)
    '    End Try
    '  End Get
    '  Set(ByVal value As ProviderInformation)
    '    Try
    '      ApplicationLogging.WriteLogEntry("Enter ProviderCatalog::Set_Item(Name)", TraceEventType.Verbose)
    '      Dim lobjProviderInformation As ProviderInformation
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjProviderInformation = CType(MyBase.Item(lintCounter), ProviderInformation)
    '        If lobjProviderInformation.Name = Name Then
    '          MyBase.Item(lintCounter) = value
    '          Exit Property
    '        End If
    '      Next
    '      Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, "ProviderCatalog::Set_Item(Name)")
    '    Finally
    '      ApplicationLogging.WriteLogEntry("Exit ProviderCatalog::Set_Item(Name)", TraceEventType.Verbose)
    '    End Try
    '  End Set
    'End Property

    '''' <summary>
    '''' Gets or sets an item in the catalog
    '''' </summary>
    '''' <param name="Index">The catalog index</param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Default Shadows Property Item(ByVal Index As Integer) As ProviderInformation
    '  Get
    '    Return MyBase.Item(Index)
    '  End Get
    '  Set(ByVal value As ProviderInformation)
    '    MyBase.Item(Index) = value
    '  End Set
    'End Property

#Region "Finalizer"

    Protected Overrides Sub Finalize()
      MyBase.Finalize()
    End Sub

#End Region

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
        Return PROVIDER_CATALOG_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ITableable Implementation"

    'Public Function ToDataTable() As DataTable Implements ITableable.ToDataTable
    '  Try

    '    ' Create New DataTable
    '    Dim lobjDataTable As New DataTable(String.Format("tbl{0}", Me.GetType.Name))

    '    ' Create columns        
    '    With lobjDataTable.Columns
    '      .Add("Name", System.Type.GetType("System.String"))

    '    End With

    '    Return lobjDataTable

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Function ToInterfaceMatrixDataTable(lpBooleanFlags As Boolean) As DataTable
      Try

        ' Create New DataTable
        Dim lobjDataTable As New DataTable("tblProviderInterfaces")
        Dim lstrFlagDataType As String

        If lpBooleanFlags Then
          lstrFlagDataType = "System.Boolean"
        Else
          lstrFlagDataType = "System.Byte"
        End If

        ' Create columns        
        With lobjDataTable.Columns
          .Add("Provider", Type.GetType("System.String"))
          .Add("Classification", Type.GetType(lstrFlagDataType))
          .Add("DocumentExport", Type.GetType(lstrFlagDataType))
          .Add("DocumentImport", Type.GetType(lstrFlagDataType))
          .Add("Explorer", Type.GetType(lstrFlagDataType))
          .Add("BasicContentServices", Type.GetType(lstrFlagDataType))
          .Add("RepositoryDiscovery", Type.GetType(lstrFlagDataType))
          .Add("AnnotationExport", Type.GetType(lstrFlagDataType))
          .Add("AnnotationImport", Type.GetType(lstrFlagDataType))
          .Add("ChoiceListExport", Type.GetType(lstrFlagDataType))
          .Add("ChoiceListImport", Type.GetType(lstrFlagDataType))
          .Add("Create", Type.GetType(lstrFlagDataType))
          .Add("Copy", Type.GetType(lstrFlagDataType))
          .Add("Delete", Type.GetType(lstrFlagDataType))
          .Add("File", Type.GetType(lstrFlagDataType))
          .Add("Rename", Type.GetType(lstrFlagDataType))
          .Add("Version", Type.GetType(lstrFlagDataType))
          .Add("FolderClassification", Type.GetType(lstrFlagDataType))
          .Add("FolderManager", Type.GetType(lstrFlagDataType))
          .Add("RecordsManager", Type.GetType(lstrFlagDataType))
          .Add("SecurityClassification", Type.GetType(lstrFlagDataType))
          .Add("UpdateProperties", Type.GetType(lstrFlagDataType))
          .Add("UpdatePermissions", Type.GetType(lstrFlagDataType))
        End With

        Sort()

        Dim lobjProvider As CProvider
        Dim lobjInterfaceDictionary As Dictionary(Of ProviderClass, Boolean)

        For Each lobjProviderInfo As ProviderInformation In Me
          lobjProvider = ContentSource.GetProvider(lobjProviderInfo.ProviderPath)

          If lobjProvider Is Nothing Then
            ApplicationLogging.WriteLogEntry(String.Format("Failed to load provider from path '{0}'.", lobjProviderInfo.ProviderPath), MethodBase.GetCurrentMethod(), TraceEventType.Warning, 18235)
            Continue For
          End If

          lobjInterfaceDictionary = lobjProvider.GetInterfaceDictionary()

          lobjDataTable.Rows.Add(lobjProviderInfo.Name,
                                          lobjInterfaceDictionary(ProviderClass.Classification),
                                          lobjInterfaceDictionary(ProviderClass.DocumentExporter),
                                          lobjInterfaceDictionary(ProviderClass.DocumentImporter),
                                          lobjInterfaceDictionary(ProviderClass.Explorer),
                                          lobjInterfaceDictionary(ProviderClass.BasicContentServices),
                                          lobjInterfaceDictionary(ProviderClass.RepositoryDiscovery),
                                          lobjInterfaceDictionary(ProviderClass.AnnotationExporter),
                                          lobjInterfaceDictionary(ProviderClass.AnnotationImporter),
                                          lobjInterfaceDictionary(ProviderClass.ChoiceListExporter),
                                          lobjInterfaceDictionary(ProviderClass.ChoiceListImporter),
                                          lobjInterfaceDictionary(ProviderClass.Create),
                                          lobjInterfaceDictionary(ProviderClass.Copy),
                                          lobjInterfaceDictionary(ProviderClass.Delete),
                                          lobjInterfaceDictionary(ProviderClass.File),
                                          lobjInterfaceDictionary(ProviderClass.Rename),
                                          lobjInterfaceDictionary(ProviderClass.Version),
                                          lobjInterfaceDictionary(ProviderClass.FolderClassification),
                                          lobjInterfaceDictionary(ProviderClass.FolderManager),
                                          lobjInterfaceDictionary(ProviderClass.RecordsManager),
                                          lobjInterfaceDictionary(ProviderClass.SecurityClassification),
                                          lobjInterfaceDictionary(ProviderClass.UpdateProperties),
                                          lobjInterfaceDictionary(ProviderClass.UpdatePermissions))

        Next

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

#Region "Provider Interface Matrix"

    'Public Function CreateProviderInterfaceMatrixSpreadsheet() As Object
    '  Try
    '    Dim lobjMatrixPackage As ExcelPackage = ToInterfaceMatrixSpreadsheet()

    '    ' Format the spreadsheet
    '    Dim lobjMainWorksheet As ExcelWorksheet = lobjMatrixPackage.Workbook.Worksheets.First
    '    lobjMainWorksheet.Tables.First.TableStyle = Table.TableStyles.Medium2

    '    Dim lintEndColumn As Integer = lobjMainWorksheet.Dimension.End.Column
    '    Dim lintEndRow As Integer = lobjMainWorksheet.Dimension.End.Row
    '    Dim lobjMatrixRange As New ExcelAddress(2, 2, lintEndRow, lintEndColumn)

    '    ' Format the matrix
    '    lobjMainWorksheet.Cells(2, 2, lintEndRow, lintEndColumn).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

    '    Dim lcfrEqualToOne As IExcelConditionalFormattingThreeIconSet(Of eExcelconditionalFormatting3IconsSetType) = lobjMainWorksheet.ConditionalFormatting.AddThreeIconSet(lobjMatrixRange, eExcelconditionalFormatting3IconsSetType.Symbols)
    '    lcfrEqualToOne.ShowValue = False
    '    ''lcfrEqualToOne.Icon1.Value = 1
    '    ''lcfrEqualToOne.Icon2.Value = 1
    '    ''lcfrEqualToOne.Icon3.Value = 100
    '    'lcfrEqualToOne.StopIfTrue = True

    '    Return lobjMatrixPackage

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function CreateProviderInterfaceMatrixSpreadsheetStream() As Stream
    '  Try
    '    Dim lobjSpreadSheet As ExcelPackage = CreateProviderInterfaceMatrixSpreadsheet()
    '    Dim lobjOutputStream As New MemoryStream
    '    lobjSpreadSheet.SaveAs(lobjOutputStream)
    '    If lobjOutputStream.CanSeek Then
    '      lobjOutputStream.Seek(0, 0)
    '    End If
    '    Return lobjOutputStream
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function CreateNamedProviderInterfaceMatrixSpreadsheetStream() As INamedStream
    '  Try
    '    Return New NamedStream(CreateProviderInterfaceMatrixSpreadsheetStream(), "InterfaceMatrixSpreadsheetStream.xlsx")
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function WriteProviderInterfaceMatrixSpreadsheet() As String
    '  Try
    '    Dim lstrOutputPath As String = Helper.CleanPath(String.Format("{0}\{1}",
    '                                             FileHelper.Instance.TempPath, "InterfaceMatrixSpreadsheetStream.xlsx"))
    '    lstrOutputPath = Helper.CleanFile(lstrOutputPath, "_")

    '    If File.Exists(lstrOutputPath) Then
    '      If Helper.IsFileLocked(lstrOutputPath) = True Then
    '        Throw New Exceptions.ItemAlreadyExistsException(lstrOutputPath, "The provider interface matrix spreadsheet already exists and is locked.")
    '      End If
    '    End If
    '    Dim lobjSpreadSheet As ExcelPackage = Me.CreateProviderInterfaceMatrixSpreadsheet
    '    If lobjSpreadSheet IsNot Nothing Then
    '      lobjSpreadSheet.SaveAs(New FileInfo(lstrOutputPath))
    '      Return lstrOutputPath
    '    Else
    '      Return String.Empty
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

#End Region

#Region "Private Methods"

    'Private Function ToInterfaceMatrixSpreadsheet() As Object
    '  Try
    '    Return ExcelHelper.CreateSpreadSheetFromDataTable("Provider Interface Matrix", ToInterfaceMatrixDataTable(False))
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

  End Class

End Namespace
