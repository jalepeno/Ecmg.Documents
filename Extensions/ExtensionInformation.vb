' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ExtensionInformation.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 4:59:23 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Extensions

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ExtensionInformation
    Implements IExtensionInformation
    Implements IXmlSerializable
    Implements IComparable

#Region "Class Variables"

    Private mstrCompanyName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrName As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mstrPath As String = String.Empty
    Private mstrProductName As String = String.Empty
    Private mblnIsValid As Boolean = True
    Private mobjLoadException As Exception = Nothing

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpName As String,
               lpDescription As String,
               lpCompanyName As String,
               lpProductName As String,
               lpPath As String)
      Me.New(lpName, lpName, lpDescription, lpCompanyName, lpProductName, lpPath)
    End Sub

    Public Sub New(lpName As String,
                   lpDisplayName As String,
                   lpDescription As String,
                   lpCompanyName As String,
                   lpProductName As String,
                   lpPath As String)
      Try
        mstrName = lpName
        mstrDisplayName = lpDisplayName
        mstrDescription = lpDescription
        mstrCompanyName = lpCompanyName
        mstrProductName = lpProductName
        mstrPath = lpPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New Text.StringBuilder
      Try

        lobjIdentifierBuilder.AppendFormat("{0} Extension: {1} - {2}({3})", Name, CompanyName, ProductName, Path)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"
    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        Return Name.CompareTo(obj.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IExtensionInformation Implementation"

    Public ReadOnly Property CompanyName As String Implements IExtensionInformation.CompanyName
      Get
        Return mstrCompanyName
      End Get
    End Property

    Public ReadOnly Property Description As String Implements IExtensionInformation.Description
      Get
        Return mstrDescription
      End Get
    End Property

    Public ReadOnly Property Name As String Implements IExtensionInformation.Name
      Get
        Return mstrName
      End Get
    End Property

    Public ReadOnly Property DisplayName As String Implements IExtensionInformation.DisplayName
      Get
        Return mstrDisplayName
      End Get
    End Property

    Public ReadOnly Property Path As String Implements IExtensionInformation.Path
      Get
        Return mstrPath
      End Get
    End Property

    Public ReadOnly Property ProductName As String Implements IExtensionInformation.ProductName
      Get
        Return mstrProductName
      End Get
    End Property

    Public Function ToXmlElementString() As String Implements IExtensionInformation.ToXmlElementString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public ReadOnly Property IsValid As Boolean Implements IExtensionInformation.IsValid
      Get
        Try
          Return mblnIsValid
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property LoadException As Exception Implements IExtensionInformation.LoadException
      Get
        Try
          Return mobjLoadException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend Sub OnLoadError(exception As Exception) Implements IExtensionInformation.OnLoadError
      Try
        If exception Is Nothing Then
          Throw New ArgumentNullException("exception")
        End If
        mobjLoadException = exception
        mblnIsValid = False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Return Nothing
    End Function

    Public Sub ReadXml(reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml

      Try

        With reader

          ' Get the Name
          mstrName = .GetAttribute("Name")

          ' Get the DisplayName
          mstrDisplayName = .GetAttribute("DisplayName")

          ' Get the Description
          mstrDescription = .GetAttribute("Description")

          ' Get the CompanyName
          mstrCompanyName = .GetAttribute("CompanyName")

          ' Get the ProductName
          mstrProductName = .GetAttribute("ProductName")

          ' Get the Path
          mstrPath = .GetAttribute("Path")

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

          ' Write the Name attribute
          .WriteAttributeString("Name", Name)

          ' Write the DisplayName attribute
          .WriteAttributeString("DisplayName", DisplayName)

          ' Write the Description attribute
          .WriteAttributeString("Description", Description)

          ' Write the CompanyName attribute
          .WriteAttributeString("CompanyName", CompanyName)

          ' Write the ProductName attribute
          .WriteAttributeString("ProductName", ProductName)

          ' Write the Path attribute
          .WriteAttributeString("Path", Path)

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