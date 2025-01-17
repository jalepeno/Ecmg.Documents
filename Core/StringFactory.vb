' ********************************************************************************
' '  Document    :  StringFactory.vb
' '  Description :  [type_description_here]
' '  Created     :  11/21/2012-1:45:53
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports System.Text
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class StringFactory

#Region "Class Variables"

    Private mstrFormatString As String = String.Empty
    Private mobjParameters As New List(Of String)

#End Region

#Region "Public Properties"

    Public Property FormatString As String
      Get
        Try
          Return mstrFormatString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrFormatString = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Parameters As List(Of String)
      Get
        Try
          Return mobjParameters
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As List(Of String))
        Try
          mobjParameters = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpFormatString As String, ParamArray lpParameters() As String)
      Try
        mstrFormatString = lpFormatString
        mobjParameters.AddRange(lpParameters)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function ToString() As String
      Try
        Dim lobjStringBuilder As New StringBuilder

        If String.IsNullOrEmpty(Me.FormatString) Then
          lobjStringBuilder.Append("FormatString not defined: ")
          If Parameters.Count = 0 Then
            lobjStringBuilder.Append("No Parameters")
          Else
            For Each lstrParameter As String In Parameters
              lobjStringBuilder.Append("{0}, ", lstrParameter)
            Next
            lobjStringBuilder.Remove(lobjStringBuilder.Length - 2, 2)
          End If
        Else
          If Parameters.Count > 0 Then
            lobjStringBuilder.AppendFormat(String.Format("{0}", FormatString), Parameters.ToArray)
          Else
            lobjStringBuilder.Append(FormatString)
          End If
        End If

        Return lobjStringBuilder.ToString

      Catch FormatEx As FormatException
        ApplicationLogging.LogException(FormatEx, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw New FormatException(DebuggerIdentifier, FormatEx)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function FromXmlString(lpXml As String) As UrlFactory
      Try
        Return Serializer.Deserialize.XmlString(lpXml, GetType(UrlFactory))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlString() As String
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlElementString() As String
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Updates the parameter values with the property values of the 
    ''' corresponding properties of the specified document.
    ''' </summary>
    ''' <param name="lpDocument" type="Ecmg.Cts.Core.Document">
    '''     <para>
    '''     The document containing the property values.
    '''     </para>
    ''' </param>
    Public Sub ResolveParametersFromDocument(lpDocument As Document)

      Try

        For i As Integer = 0 To Parameters.Count - 1

          Dim lstrParam As String = Parameters(i)

          If (lpDocument.PropertyExists(PropertyScope.BothDocumentAndVersionProperties, lstrParam)) Then
            Dim lobjProperty As ECMProperty = lpDocument.GetProperty(lstrParam)
            If lobjProperty.Cardinality = Cardinality.ecmSingleValued Then
              Parameters(i) = lobjProperty.Value
            Else
              Throw New Exceptions.InvalidCardinalityException(
                String.Format("Could not resolve parameter from document. The property '{0}' is multi-valued.",
                              lstrParam), Cardinality.ecmSingleValued, lobjProperty)
            End If
          Else
            Throw New Exceptions.PropertyDoesNotExistException(
              String.Format("Could not resolve parameter from document. The property '{0}' does not exist.",
                            lstrParam), lstrParam)
          End If

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try
        If String.IsNullOrEmpty(Me.FormatString) Then
          lobjIdentifierBuilder.Append("FormatString not defined: ")
          If Parameters.Count = 0 Then
            lobjIdentifierBuilder.Append("No Parameters")
          Else
            lobjIdentifierBuilder.Append(GetDelimitedParameterString(", "))
          End If
        Else
          Dim lintExpectedParameterCount As Integer = FormatString.Split("{").Length - 1
          If lintExpectedParameterCount <> Parameters.Count Then
            lobjIdentifierBuilder.AppendFormat("{0} (Expected parameter count is {1}: parameters provided<{2}>)",
                                               FormatString, lintExpectedParameterCount, GetDelimitedParameterString(", "))
          End If
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function GetDelimitedParameterString(lpDelimiter As String) As String
      Try
        Dim lobjStringBuilder As New StringBuilder

        If Parameters.Count = 0 Then
          Return String.Empty
        End If

        For Each lstrParameter As String In Parameters
          lobjStringBuilder.AppendFormat("{0}{1}", lstrParameter, lpDelimiter)
        Next
        lobjStringBuilder.Remove(lobjStringBuilder.Length - lpDelimiter.Length, lpDelimiter.Length)

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace