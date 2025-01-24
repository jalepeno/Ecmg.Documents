' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ExtensionCatalog.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 4:55:48 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Extensions

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ExtensionCatalog
    Implements IXmlSerializable
    Implements INotifyCollectionChanged

#Region "Class Constants"

    Private Const FILE_NAME As String = "ExtensionCatalog.xml"

#End Region

#Region "Class Variables"

    Private WithEvents mobjExtensionEntries As New ExtensionEntries

    Private Shared mobjInstance As ExtensionCatalog
    Private Shared mintReferenceCount As Integer
    Private Shared mstrCatalogPath As String = String.Empty

    Private mobjExtensionTypes As List(Of Type) = Nothing
    Private mobjAvailableExtensions As IDictionary(Of String, Type) = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property Extensions As ExtensionEntries
      Get
        Return mobjExtensionEntries
      End Get
    End Property

    Public ReadOnly Property AvailableExtensions() As IDictionary(Of String, Type)
      Get
        Try
          If mobjAvailableExtensions Is Nothing Then
            mobjAvailableExtensions = GetAvailableExtensions()
          End If
          Return mobjAvailableExtensions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Private ReadOnly Property Types As List(Of Type)
      Get
        Try
          If mobjExtensionTypes Is Nothing Then
            mobjExtensionTypes = GetAllExtensionTypes()
          End If
          Return mobjExtensionTypes
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shared Property CatalogPath As String
      Get
        If mobjInstance Is Nothing Then
          mobjInstance = GetCurrentCatalog()
        End If
        Return mstrCatalogPath
      End Get
      Private Set(value As String)
        mstrCatalogPath = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    Public Sub Add(lpExtension As IExtensionInformation)
      Try
        If Me.Extensions.Contains(lpExtension.Name) = True Then
          Throw New Exceptions.ItemAlreadyExistsException(lpExtension,
            String.Format("Can't add extension '{0}'.  An extension by that name already exists in the catalog.", lpExtension.Name))
        Else
          Me.Extensions.Add(lpExtension)
          Me.Extensions.Sort()
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Add(lpName As String, lpDisplayName As String, lpDescription As String, lpCompanyName As String, lpProductName As String, lpPath As String)
      Try
        ' Check the path
        If File.Exists(lpPath) = False Then
          Throw New Exceptions.InvalidPathException(
            String.Format("Unable to add extension, the specified path '{0}' is invalid.", lpPath), lpPath)
        End If

        Add(New ExtensionInformation(lpName, lpDisplayName, lpDescription, lpCompanyName, lpProductName, lpPath))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Remove(lpName As String)
      Try
        If Me.Extensions.Contains(lpName) = True Then
          Me.Extensions.Remove(Me.Extensions(lpName))
        Else
          Throw New Exceptions.ItemDoesNotExistException(lpName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Singleton Support"

    Public Shared Function Instance() As ExtensionCatalog
      Try
        If mobjInstance Is Nothing Then
          mobjInstance = GetCurrentCatalog()
        End If
        mintReferenceCount += 1
        Return mobjInstance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

#End Region

#Region "Private Methods"

    Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New Text.StringBuilder
      Try

        lobjIdentifierBuilder.Append("Extension Catalog")

        Select Case Extensions.Count
          Case 0
            lobjIdentifierBuilder.Append(": No Extensions")
          Case 1
            lobjIdentifierBuilder.Append(": 1 Extension")
          Case Else
            lobjIdentifierBuilder.AppendFormat(": {0} Extensions", Extensions.Count)
        End Select

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetCurrentCatalog() As ExtensionCatalog
      Try

        mstrCatalogPath = Helper.CleanPath(String.Format("{0}\{1}", FileHelper.Instance.ConfigPath, FILE_NAME))

        If File.Exists(mstrCatalogPath) Then
          Try
            Return Serializer.Deserialize.XmlFile(mstrCatalogPath, GetType(ExtensionCatalog))
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            Return New ExtensionCatalog
          End Try
        Else
          Return New ExtensionCatalog
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetAllExtensionTypes() As List(Of Type)

      Dim lobjOperationList As New List(Of String)
      Dim lobjExtensionType As Type = Nothing
      Dim lobjAssembly As Reflection.Assembly = Nothing
      Dim lobjTypes As Object = Nothing
      Dim lobjExtensionTypes As New List(Of Type)
      Dim lobjExtension As IExtension = Nothing

      Try

        lobjExtensionType = GetType(IExtension)

        For Each lobjExtensionInfo As IExtensionInformation In Extensions

          Try
            lobjAssembly = Assembly.LoadFrom(lobjExtensionInfo.Path)
          Catch ex As Exception
            ApplicationLogging.WriteLogEntry(String.Format("Failed to load assembly from file: '{0}'.", lobjExtensionInfo.Path),
                                             Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Error, 198234)
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            lobjExtensionInfo.OnLoadError(ex)
            Continue For
          End Try

          Try
            lobjTypes = lobjAssembly.GetTypes.Where(Function(t) lobjExtensionType.IsAssignableFrom(t))
          Catch ex As Exception
            ApplicationLogging.WriteLogEntry(String.Format("Failed to get types from file: '{0}'.", lobjExtensionInfo.Path),
                                             Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Error, 198235)

            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            lobjExtensionInfo.OnLoadError(ex)
            Continue For
          End Try

          lobjExtension = Nothing

          For Each lobjType As Type In lobjTypes

            If lobjType.IsAbstract Then
              Continue For
            End If

            If lobjType.IsInterface Then
              Continue For
            End If

            lobjExtension = Activator.CreateInstance(lobjType)

            If lobjExtension IsNot Nothing Then
              lobjExtensionTypes.Add(lobjType)
            End If

            lobjExtension = Nothing

          Next

        Next

        Return lobjExtensionTypes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ' Clean up references
        lobjAssembly = Nothing
        lobjTypes = Nothing
        lobjExtensionType = Nothing
      End Try
    End Function

    Private Function GetAvailableExtensions() As IDictionary(Of String, Type)

      Dim lobjExtensionList As New Dictionary(Of String, Type)
      'Dim lobjExtensionInformation As IExtensionInformation = Nothing
      Dim lobjExtensionInformation As Object = Nothing

      Try

        For Each lobjType As Type In Me.Types

          lobjExtensionInformation = Activator.CreateInstance(lobjType)

          If lobjExtensionInformation IsNot Nothing AndAlso lobjExtensionList.ContainsKey(lobjExtensionInformation.Name) = False Then
            lobjExtensionList.Add(lobjExtensionInformation.Name, lobjType)
          End If

          lobjExtensionInformation = Nothing

        Next

        Return lobjExtensionList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Return Nothing
    End Function

    Public Sub ReadXml(reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml

      Dim lobjXmlDocument As New XmlDocument
      Dim lstrName As String = String.Empty
      Dim lstrDisplayName As String = String.Empty
      Dim lstrDescription As String = String.Empty
      Dim lstrCompanyName As String = String.Empty
      Dim lstrProductName As String = String.Empty
      Dim lstrPath As String = String.Empty

      Try

        With reader

          lobjXmlDocument.Load(reader)

          With lobjXmlDocument

            ' Read the ExtensionInformation elements
            Dim lobjExtensionNodes As XmlNodeList = .SelectNodes("//Extensions/ExtensionInformation")

            For Each lobjExtensionNode As XmlNode In lobjExtensionNodes

              ' Read the name
              If lobjExtensionNode.Attributes("Name") IsNot Nothing Then
                lstrName = lobjExtensionNode.Attributes("Name").Value
              Else
                Throw New InvalidOperationException("The extension xml has no Name attribute.")
              End If

              ' Read the display name
              If lobjExtensionNode.Attributes("DisplayName") IsNot Nothing Then
                lstrDisplayName = lobjExtensionNode.Attributes("DisplayName").Value
              Else
                lstrDisplayName = lstrName
              End If

              ' Read the description
              If lobjExtensionNode.Attributes("Description") IsNot Nothing Then
                lstrDescription = lobjExtensionNode.Attributes("Description").Value
              End If

              ' Read the company name
              If lobjExtensionNode.Attributes("CompanyName") IsNot Nothing Then
                lstrCompanyName = lobjExtensionNode.Attributes("CompanyName").Value
              Else
                Throw New InvalidOperationException("The extension xml has no CompanyName attribute.")
              End If

              ' Read the product name
              If lobjExtensionNode.Attributes("ProductName") IsNot Nothing Then
                lstrProductName = lobjExtensionNode.Attributes("ProductName").Value
              Else
                Throw New InvalidOperationException("The extension xml has no ProductName attribute.")
              End If

              ' Read the path
              If lobjExtensionNode.Attributes("Path") IsNot Nothing Then
                lstrPath = lobjExtensionNode.Attributes("Path").Value
              Else
                Throw New InvalidOperationException("The extension xml has no Path attribute.")
              End If

              ' Add the extension to the catalog
              Me.Extensions.Add(New ExtensionInformation(lstrName, lstrDisplayName, lstrDescription, lstrCompanyName, lstrProductName, lstrPath))

            Next

          End With

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          If Extensions.Count > 0 Then
            .WriteStartElement("Extensions")
            For Each lobjExtension As IExtensionInformation In Extensions
              ' Write the Extension element
              .WriteRaw(lobjExtension.ToXmlElementString)
            Next
            .WriteEndElement()
          Else
            .WriteElementString("Extensions", Nothing)
          End If
        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "INotifyCollectionChanged Implementation"

    Public Event CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Implements INotifyCollectionChanged.CollectionChanged

    Private Sub mobjExtensionEntries_CollectionChanged(sender As Object, e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles mobjExtensionEntries.CollectionChanged
      Try
        If Not Helper.CallStackContainsMethodName("ReadXml") Then
          Serializer.Serialize.XmlFile(Me, mstrCatalogPath)
          Helper.FormatXmlFile(mstrCatalogPath)
          mobjInstance = GetCurrentCatalog()
          RaiseEvent CollectionChanged(sender, e)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace