'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  TransformationConfiguration.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 2:52:51 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Transformations

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TransformationConfiguration
    Implements ITransformationConfiguration
    Implements IXmlSerializable

#Region "Class Constants"

    Private Const XML_HEADER As String = "<?xml version=""1.0"" encoding=""utf-16""?>"

#End Region

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mstrDescription As String = String.Empty

    Private WithEvents mobjActions As IActionItems = New ActionItems

#End Region

#Region "Public Properties"

    Public Property Name As String Implements ITransformationConfiguration.Name
      Get
        Try
          Return mstrName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property DisplayName As String Implements ITransformationConfiguration.DisplayName
      Get
        Try
          If String.IsNullOrEmpty(mstrDisplayName) Then
            mstrDisplayName = Helper.CreateDisplayName(Me.Name)
          End If
          Return mstrDisplayName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property Description As String Implements ITransformationConfiguration.Description
      Get
        Try
          Return mstrDescription
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrDescription = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Actions As IActionItems Implements ITransformationConfiguration.Actions
      Get
        Try
          Return mobjActions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IActionItems)
        Try
          mobjActions = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Public Functions"

    Public Function ToTransformation() As Transformation Implements ITransformationConfiguration.ToTransformation
      Try

        Return New Transformation(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlString() As String Implements ITransformationConfiguration.ToXmlString
      Try
        Return Helper.FormatXmlString(Serializer.Serialize.XmlString(Me))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Friend Function DebuggerIdentifier() As String
      Try
        If Not String.IsNullOrEmpty(Name) Then
          Return String.Format("{0}: {1} Action(s)", Name, Actions.Count)
        Else
          Return String.Format("Unnamed Transformation: {0} Action(s)", Actions.Count)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Shared Function GetActions(ByVal lpActionsNode As XmlNode) As IActionItems
      Try

        Dim lobjActions As IActionItems = New ActionItems
        Dim lobjAction As IActionItem = Nothing

        For Each lobjActionNode As XmlNode In lpActionsNode.ChildNodes

          ' Read the action type
          ' This assumes that the action type is the first attribute
          If lobjActionNode.Attributes Is Nothing OrElse lobjActionNode.Attributes.Count = 0 Then
            Throw New InvalidOperationException("The action xml has no type attribute.")
          End If

          lobjAction = ActionFactory.Create(lobjActionNode, CultureInfo.CurrentCulture.Name)

          lobjActions.Add(New ActionItem(lobjAction))

        Next

        Return lobjActions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetSchema() As Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
      Try
        ' As per the Microsoft guidelines this is not implemented
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub ReadXml(reader As Xml.XmlReader) Implements IXmlSerializable.ReadXml
      Try
        Dim lobjProcessXmlBuilder As New StringBuilder
        Dim lobjXmlDocument As New XmlDocument

        ' We were having problems reading when loading as part of a larger object such as JobConfiguration.  
        ' Recreating the xml string seems to resolve the issue.
        lobjProcessXmlBuilder.AppendLine(XML_HEADER)
        lobjProcessXmlBuilder.Append(reader.ReadOuterXml)

        lobjXmlDocument.LoadXml(lobjProcessXmlBuilder.ToString)

        With lobjXmlDocument

          ' Get the name
          mstrName = .DocumentElement.GetAttribute("Name")

          ' Get the display name
          mstrDisplayName = .DocumentElement.GetAttribute("DisplayName")

          ' Get the description
          mstrDescription = .DocumentElement.GetAttribute("Description")

          ' Read the ActionItem elements
          Dim lobjOperationsNode As XmlNode = .SelectSingleNode("//Actions")

          Actions.AddRange(GetActions(lobjOperationsNode))

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(writer As Xml.XmlWriter) Implements IXmlSerializable.WriteXml
      Try
        With writer

          ' Write the Name attribute
          .WriteAttributeString("Name", Me.Name)

          ' Write the DisplayName attribute
          .WriteAttributeString("DisplayName", Me.Name)

          ' Write the Description attribute
          .WriteAttributeString("Description", Me.Description)

          .WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
          .WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")

          ' Open the Actions Element
          .WriteStartElement("Actions")
          ' Write out the parameters
          If Me.Actions IsNot Nothing Then
            For Each lobjItem As IActionItem In Me.Actions
              ' Write the Parameter element
              .WriteRaw(lobjItem.ToXmlElementString)
            Next
          End If
          ' End the Actions element
          .WriteEndElement()

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