'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ActionItem.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 11:06:38 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core

  <TypeConverter(GetType(ExpandableObjectConverter)),
  DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ActionItem
    Implements IActionItem
    Implements IXmlSerializable

#Region "Class Variables"

    Private mstrDescription As String = String.Empty
    Private ReadOnly mstrDisplayName As String = String.Empty
    Private ReadOnly mstrName As String = String.Empty
    Private WithEvents MobjParameters As IParameters = New Parameters

#End Region

#Region "Public Properties"

    Public ReadOnly Property Name As String Implements IActionItem.Name
      Get
        Try
          Return mstrName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property Description As String Implements IActionItem.Description
      Get
        Try
          Return mstrDescription
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property DisplayName As String Implements IActionItem.DisplayName
      Get
        Try
          Return mstrDisplayName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property Parameters As Core.IParameters Implements IActionItem.Parameters
      Get
        Return MobjParameters
      End Get
      Set(value As Core.IParameters)
        Try
          MobjParameters = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        'If Parameters.Count = 0 Then
        '  Parameters = GetDefaultParameters()
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpActionItem As IActionItem)
      Try
        mstrName = lpActionItem.Name
        mstrDescription = lpActionItem.Description
        mstrDisplayName = lpActionItem.DisplayName
        MobjParameters = lpActionItem.Parameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
      Try
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetBooleanParameterValue(lpParameterName As String, lpDefaultValue As Object) As Boolean Implements IActionItem.GetBooleanParameterValue
      Try
        Return GetBooleanParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetEnumParameterValue(lpParameterName As String, lpEnumType As Type, lpDefaultValue As Object) As [Enum] Implements IActionItem.GetEnumParameterValue
      Try
        Return GetEnumParameterValue(Me, lpParameterName, lpEnumType, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetStringParameterValue(lpParameterName As String, lpDefaultValue As Object) As String Implements IActionItem.GetStringParameterValue
      Try
        Return GetStringParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetParameterValue(lpParameterName As String, lpDefaultValue As Object) As Object Implements IActionItem.GetParameterValue
      Try
        Return GetParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SetDescription(ByVal lpDescription As String) Implements IActionItem.SetDescription
      Try
        mstrDescription = lpDescription
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Shared Methods"

    Public Shared Function GetParameterValue(ByVal lpActionable As IActionItem, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object
      Try
        If lpActionable.Parameters.Contains(lpParameterName) Then
          ' Return lpOperable.Parameters.Item(lpParameterName).Value
          Dim lobjValue As Object = lpActionable.Parameters.Item(lpParameterName).Value
          If lobjValue IsNot Nothing AndAlso TypeOf lobjValue Is String Then
            Dim lstrValue As String = lobjValue.ToString
            Dim lobjParamNameValuePair As INameValuePair = Parameter.GetInlineParameter(lstrValue)

            If lobjParamNameValuePair Is Nothing Then
              Return lstrValue
            End If

            Dim lstrRequestedParameter As String = String.Empty

            Select Case lobjParamNameValuePair.Name.ToLower
              Case "parameter"
                lstrRequestedParameter = lpActionable.GetStringParameterValue(lobjParamNameValuePair.Value, String.Empty)
              Case Else
                Return lstrValue
            End Select

            If String.IsNullOrEmpty(lstrRequestedParameter) Then
              Return lstrValue
            End If

            Dim lobjRegex As New Regex(Parameter.PARAMETER_EXPRESSION,
                RegexOptions.CultureInvariant Or RegexOptions.Compiled)

            ' Split the InputText wherever the regex matches
            Dim lstrResults As String() = lobjRegex.Split(lstrValue)

            ' Test to see if there is a match in the InputText
            Dim lblnIsMatch As Boolean = lobjRegex.IsMatch(lstrValue)

            If lblnIsMatch Then
              Dim lintPrefixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Prefix")
              Dim lintSuffixGroupNumber As Integer = lobjRegex.GroupNumberFromName("Suffix")
              Dim lobjStringBuilder As New StringBuilder

              If lintPrefixGroupNumber > 0 Then
                lobjStringBuilder.Append(lstrResults(lintPrefixGroupNumber))
              End If
              lobjStringBuilder.Append(lstrRequestedParameter)
              If lintSuffixGroupNumber > 0 Then
                lobjStringBuilder.Append(lstrResults(lintSuffixGroupNumber))
              End If

              Return lobjStringBuilder.ToString

            Else
              Return lstrValue
            End If
          Else
            Return lobjValue
          End If
        Else
          Return lpDefaultValue
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetBooleanParameterValue(ByVal lpActionable As IActionItem, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean
      Try
        Return CBool(GetParameterValue(lpActionable, lpParameterName, lpDefaultValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetEnumParameterValue(ByVal lpActionable As IActionItem, ByVal lpParameterName As String, lpEnumType As Type, ByVal lpDefaultValue As Object) As [Enum]
      Try
        Return [Enum].Parse(lpEnumType, GetStringParameterValue(lpActionable, lpParameterName, lpDefaultValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetStringParameterValue(ByVal lpActionable As IActionItem, ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As String
      Try
        Return CStr(GetParameterValue(lpActionable, lpParameterName, lpDefaultValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub ReadOperableXml(ByRef sender As IActionItem, ByVal reader As System.Xml.XmlReader)

      Try

        If reader.IsEmptyElement Then
          reader.Read()
          Exit Sub
        End If

        Dim lstrCurrentElementName As String = String.Empty

        sender.SetDescription(reader.GetAttribute("Description"))

        'Boolean.TryParse(reader.GetAttribute("LogResult"), sender.LogResult)

        'If TypeOf sender Is IOperation Then
        '  Dim lstrScope As String = reader.GetAttribute("Scope")
        '  If Not String.IsNullOrEmpty(lstrScope) Then
        '    CType(sender, IOperation).Scope = CType([Enum].Parse(GetType(OperationScope), reader.GetAttribute("Scope")), OperationScope)
        '  End If
        'End If

        Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name.EndsWith("Action")
          If reader.NodeType = XmlNodeType.Element Then
            lstrCurrentElementName = reader.Name
          Else
            Select Case lstrCurrentElementName
              Case "Parameters"
                ' Skip to the next line
              Case "Parameter"
                ' TODO: Add the parameter
              Case "Values"
                ' Skip to the next line
              Case "Value"

              Case "TrueOperations"
                ' TODO: Add the TrueOperations

              Case "FalseOperations"
                ' TODO: Add the FalseOperations

              Case "RunBeforeBegin"
                ' TODO: Implement this...

              Case "RunAfterComplete"
                ' TODO: Implement this...

              Case "RunOnFailure"
                ' TODO: Implement this...

            End Select
          End If
        Loop

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub WriteActionXml(ByVal sender As IActionItem, ByVal writer As System.Xml.XmlWriter)
      Try

        With writer

          ' Write the Name attribute
          .WriteAttributeString("Name", sender.Name)

          ' Write the Description attribute
          .WriteAttributeString("Description", sender.Description)

          ' Write the Parameters
          ' Open the Parameters Element
          .WriteStartElement("Parameters")

          If sender.Parameters IsNot Nothing Then
            For Each lobjParameter As IParameter In sender.Parameters
              ' Write the Parameter element
              .WriteRaw(lobjParameter.ToXmlString)
            Next
          End If

          ' End the Parameters element
          .WriteEndElement()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlString() As String Implements IActionItem.ToXmlElementString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If String.IsNullOrEmpty(Name) Then
          Return Me.GetType.Name
        Else
          Return Name
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICloneable Implementation"

    Public Overloads Function Clone() As Object
      Try
        Return New Parameter(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As Schema.XmlSchema Implements IXmlSerializable.GetSchema
      Try
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub ReadXml(reader As XmlReader) Implements IXmlSerializable.ReadXml

    End Sub

    Public Sub WriteXml(writer As XmlWriter) Implements IXmlSerializable.WriteXml
      Try
        WriteActionXml(Me, writer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function GetHashCode() As Integer
      Throw New NotImplementedException()
    End Function

#End Region

  End Class

End Namespace