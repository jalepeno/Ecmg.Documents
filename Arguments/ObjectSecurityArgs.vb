' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectSecurityArgs.vb
'  Description :  [type_description_here]
'  Created     :  3/21/2012 3:42:51 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Security
Imports Documents.Utilities

#End Region

Namespace Arguments

  <XmlInclude(GetType(DocumentSecurityArgs))>
  <XmlInclude(GetType(FolderSecurityArgs))>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ObjectSecurityArgs
    'Implements IXmlSerializable

#Region "Class Enumerations"

    Public Enum UpdateMode
      Replace = 0
      Append = 1
      Update = 2
    End Enum

#End Region

#Region "Class Variables"

    Private mstrObjectID As String = String.Empty
    Private mobjPermissions As IPermissions = New Permissions
    Private menuMode As UpdateMode = UpdateMode.Replace
    Private mobjTag As Object

#End Region

#Region "Public Properties"

    <System.Xml.Serialization.XmlIgnore()>
    Public Property Tag As Object
      Get
        Return mobjTag
      End Get
      Set(value As Object)
        mobjTag = value
      End Set
    End Property

    ''' <summary>
    ''' Identifies the object to be updated
    ''' </summary>
    Public Property ObjectID() As String
      Get
        Return mstrObjectID
      End Get
      Set(ByVal value As String)
        mstrObjectID = value
      End Set
    End Property

    ''' <summary>
    ''' The permissions to update in the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This will replace all existing permissions.</remarks>
    Public Property Permissions() As Permissions
      Get
        Return mobjPermissions
      End Get
      Set(ByVal value As Permissions)
        mobjPermissions = value
      End Set
    End Property

    Public Property Mode As UpdateMode
      Get
        Return menuMode
      End Get
      Set(value As UpdateMode)
        menuMode = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a new instance of ObjectSecurityArgs
    ''' </summary>
    ''' <param name="lpObjectId">The object identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String)
      Me.New(lpObjectId, New Permissions)
    End Sub

    ''' <summary>
    ''' Creates a new instance of ObjectSecurityArgs
    ''' </summary>
    ''' <param name="lpObjectId">The object identifier</param>
    ''' <param name="lpPermissions">A IPermissions object containing the permissions to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String, ByVal lpPermissions As IPermissions)

      Try

        ObjectID = lpObjectId
        Permissions = lpPermissions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

#End Region

#Region "Protected Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        Select Case Permissions.Count
          Case 1
            lobjIdentifierBuilder.AppendFormat("{0}: {1} {2} Permission",
              Me.GetType.Name, Mode.ToString, Permissions.Count)
          Case Else
            lobjIdentifierBuilder.AppendFormat("{0}: {1} {2} Permissions",
              Me.GetType.Name, Mode.ToString, Permissions.Count)
        End Select

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

    '#Region "IXmlSerializable Implementation"

    '    Public Function GetSchema() As Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
    '      Try
    '        Return Nothing
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        '  Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Function

    '    Public Sub ReadXml(reader As Xml.XmlReader) Implements IXmlSerializable.ReadXml
    '      Try

    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        '  Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Sub WriteXml(writer As Xml.XmlWriter) Implements IXmlSerializable.WriteXml
    '      Try
    '        With writer
    '          ' Write the ObjectID element
    '          .WriteElementString("ObjectID", Me.ObjectID)

    '          ' Open the Permissions Element
    '          .WriteStartElement("Permissions")
    '          ' Write out the Permissions
    '          If Me.Permissions IsNot Nothing Then
    '            For Each lobjPermission As IPermission In Me.Permissions
    '              ' Write the Permission element
    '              .WriteRaw(Serializer.Serialize.XmlElementString(lobjPermission))
    '            Next
    '          End If
    '          ' End the Permissions element
    '          .WriteEndElement()

    '        End With
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        '  Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '#End Region

  End Class

End Namespace