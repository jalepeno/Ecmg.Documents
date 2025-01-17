'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Providers
  ''' <summary>
  ''' Identifies the standard properties of a provider
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class ProviderSystem
    Implements ISerialize
    Implements ITableable

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const PROVIDER_SYSTEM_FILE_EXTENSION As String = "psf"

#End Region

#Region "Class Variables"

    Private mstrName As String
    Private mstrType As String
    Private mstrCompany As String
    Private mstrProductName As String
    Private mstrProductVersion As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpType As String, ByVal lpCompany As String,
                   ByVal lpProductName As String, ByVal lpProductVersion As String)

      mstrName = lpName
      mstrType = lpType
      mstrCompany = lpCompany
      mstrProductName = lpProductName
      mstrProductVersion = lpProductVersion

    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The provider name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' The specific type of provider for the product.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Some products may expose more than one API.  
    ''' This property is usually used to describe the particular product/API combination.</remarks>
    Public Property Type() As String
      Get
        Return mstrType
      End Get
      Set(ByVal value As String)
        mstrType = value
      End Set
    End Property

    ''' <summary>
    ''' The company that creates or produces the product to which the provider connects to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Company() As String
      Get
        Return mstrCompany
      End Get
      Set(ByVal value As String)
        mstrCompany = value
      End Set
    End Property

    ''' <summary>
    ''' The name of the product to which the provider connects to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProductName() As String
      Get
        Return mstrProductName
      End Get
      Set(ByVal value As String)
        mstrProductName = value
      End Set
    End Property

    ''' <summary>
    ''' If applicable, the version of the product for which the provider targets.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProductVersion() As String
      Get
        Return mstrProductVersion
      End Get
      Set(ByVal value As String)
        mstrProductVersion = value
      End Set
    End Property

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
        Return PROVIDER_SYSTEM_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Return Serializer.Serialize.Xml(Me)
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

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      If lpFileExtension.Length = 0 Then
        ' No override was provided
        If lpFilePath.EndsWith("cps") = False Then
          lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & "cps"
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

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Return Serializer.Serialize.XmlElementString(Me)
    End Function

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As System.Data.DataTable Implements ITableable.ToDataTable
      Try
        Dim lobjDataTable As New DataTable(String.Format("tbl{0}", Me.GetType.Name))

        lobjDataTable.Columns.Add("PropertyName")
        lobjDataTable.Columns.Add("PropertyValue")
        lobjDataTable.Rows.Add("Name", Me.Name)
        lobjDataTable.Rows.Add("Type", Me.Type)
        lobjDataTable.Rows.Add("Company", Me.Company)
        lobjDataTable.Rows.Add("ProductName", Me.ProductName)
        lobjDataTable.Rows.Add("ProductVersion", Me.ProductVersion)

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