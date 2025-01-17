' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  AccessRights.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:57:02 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"



Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities


#End Region

Namespace Security

  Public Class AccessRights
    Inherits CCollection(Of AccessRight)
    Implements IAccessRights
    Implements IXmlSerializable

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IAccessMask)

#End Region

#Region "IAccessRights Implementation"

    Public Overloads Function Contains(name As String) As Boolean Implements IAccessRights.Contains
      Try
        Return MyBase.Contains(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Property Item(name As String) As IAccessMask Implements IAccessRights.Item
      Get
        Try
          Return MyBase.Item(name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IAccessMask)
        Try
          MyBase.Item(name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Shadows ReadOnly Property Item(maskValue As Integer) As IAccessMask Implements IAccessRights.Item
      Get
        Try
          Return MyBase.FirstOrDefault(Function(IProperty) IProperty.Value = maskValue)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads Sub Add(item As IAccessMask) Implements System.Collections.Generic.ICollection(Of IAccessMask).Add
      Try
        MyBase.Add(CType(item, AccessRight))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Clear() Implements System.Collections.Generic.ICollection(Of IAccessMask).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IAccessMask) As Boolean Implements System.Collections.Generic.ICollection(Of IAccessMask).Contains
      Try
        Return Contains(CType(item, AccessRight))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IAccessMask, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IAccessMask).CopyTo
      Try
        MyBase.CopyTo(CType(array, IEnumeratedValue()), arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IAccessMask).Count
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shadows ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IAccessMask).IsReadOnly
      Get
        Try
          Return MyBase.IsReadOnly
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads Function Remove(item As IAccessMask) As Boolean Implements System.Collections.Generic.ICollection(Of IAccessMask).Remove
      Try
        Return Remove(CType(item, AccessRight))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IAccessMask) Implements System.Collections.Generic.IEnumerable(Of IAccessMask).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IAccessMask)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IAccessMask)(Me.ToArray, GetType(AccessRight), GetType(IAccessMask))
          End If
          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Try
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub ReadXml(reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try

        Dim lstrRootName As String = String.Empty
        Dim lstrClassName As String = String.Empty
        Dim lstrCurrentElementName As String = String.Empty
        Dim lobjItem As IAccessMask = Nothing

        If reader.IsEmptyElement Then
          reader.Read()
          Exit Sub
        End If

        lstrRootName = reader.Name

        Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name = lstrRootName
          If reader.NodeType = XmlNodeType.Element Then
            lstrCurrentElementName = reader.Name
          ElseIf reader.NodeType = XmlNodeType.Text Then

            Select Case lstrCurrentElementName
              Case "AccessRight"
                If lobjItem IsNot Nothing Then
                  Add(lobjItem)
                End If

                ' We are at a new item
                lstrClassName = lstrCurrentElementName

                lobjItem = New AccessRight

              Case "AccessLevel"
                If lobjItem IsNot Nothing Then
                  Add(lobjItem)
                End If

                ' We are at a new item
                lstrClassName = lstrCurrentElementName

                lobjItem = New AccessLevel

              Case "Name"
                lobjItem.Name = reader.Value

              Case "Value"
                lobjItem.Value = reader.Value

            End Select

          End If
          reader.Read()
          Do Until reader.NodeType <> XmlNodeType.Whitespace
            reader.Read()
          Loop
        Loop

        lstrClassName = reader.Name


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          For Each lobjItem As IAccessMask In Items

            If TypeOf lobjItem Is IAccessLevel Then
              .WriteStartElement("AccessLevel")
            ElseIf TypeOf lobjItem Is IAccessRight Then
              .WriteStartElement("AccessRight")
            End If

            ' Write the Name attribute
            .WriteAttributeString("Name", lobjItem.Name)

            ' Write the Value attribute
            .WriteAttributeString("Value", lobjItem.Value)
            'If lobjItem.Value.HasValue Then
            '  .WriteAttributeString("Value", lobjItem.Value)
            'Else
            '  .WriteAttributeString("Value", String.Empty)
            'End If

            ' Write the Type attribute
            .WriteAttributeString("Type", lobjItem.Type.ToString)

            .WriteStartElement("Permissions")

            If lobjItem.PermissionList.Count > 0 Then
              .WriteRaw(lobjItem.PermissionList.ToDelimitedList)
            End If

            .WriteEndElement() '"Permissions"

            .WriteEndElement()

          Next

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