'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Scripting

  ''' <summary>
  ''' In instance of a Cts Script
  ''' </summary>
  ''' <remarks></remarks>
  <XmlRoot("ctscript")>
  Public Class Script

#Region "Class Variables"

    Private mobjMethods As New Methods

#End Region

#Region "Public Properties"

    Public Property Methods() As Methods
      Get
        Return mobjMethods
      End Get
      Set(ByVal value As Methods)
        mobjMethods = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Evaluates whether value is a valid script;
    ''' </summary>
    ''' <param name="lpValue">String representation of a content transformation script</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsCtScript(ByVal lpValue As Object) As Boolean

      Try
        'Dim lobjScript As Script

        'lobjScript = Serializer.Deserialize.SoapString(lpValue, GetType(Script))

        'Return True

        If String.IsNullOrEmpty(lpValue) Then
          Return False
        End If

        Dim lstrValue As String = lpValue.ToString

        If lstrValue.StartsWith("%") AndAlso lstrValue.EndsWith("%") Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception

        ApplicationLogging.WriteLogEntry(String.Format("The script '{0}' is not a valid content transformation script: {1}", lpValue, ex),
                             TraceEventType.Information)
        Return False
      End Try

    End Function

    Public Shared Function GetValue(ByVal lpScript As String) As Object
      Try
        Return GetValue(lpScript, Core.PropertyType.ecmString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Executes the provided script and returns the script value
    ''' </summary>
    ''' <param name="lpScript">The script to be executed (expressed as "%script%")</param>
    ''' <param name="lpTargetPropertyType">The destination property type 
    ''' (using a valid Ecmg.Cts.Core.PropertyType enumeration value)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetValue(ByVal lpScript As String,
                                    ByVal lpTargetPropertyType As Integer) As Object
      Try

        ' Make sure it is a valid script
        If IsCtScript(lpScript) = False Then
          Throw New InvalidScriptException("An invalid script was provided", lpScript)
        End If

        Dim lstrScript As String = lpScript.Trim("%")

        If lstrScript.ToLower.StartsWith("now") OrElse lstrScript.ToLower = "today" Then
          Return ScriptMethods.Now(lpTargetPropertyType, lstrScript)
        End If

        Return lstrScript

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Overrides Function ToString() As String
    '  Try
    '    Return Serializer.Serialize.XmlString(Me)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Takes a content tranformation script as an XML string parameter 
    '''' and loads it into the current script instance
    '''' </summary>
    '''' <param name="lpScript">Content Transformation Script XML String</param>
    '''' <remarks></remarks>
    'Public Sub Load(ByVal lpScript As String)
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    '''' <summary>
    '''' Factory method for creating a new script object from a content transformation script xml string
    '''' </summary>
    '''' <param name="lpScript">Content Transformation Script XML String</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function Create(ByVal lpScript As String) As Script
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function Execute() As Object
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  Finally

    '  End Try
    'End Function

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpMethods As Methods)
      Methods = lpMethods
    End Sub

#End Region

  End Class

  Public Class ScriptResult
    Inherits Core.ActionResult

#Region "Class Variables"

    Private mobjScript As Script

#End Region

#Region "Public Properties"

    Public Property Script() As Script
      Get
        Return mobjScript
      End Get
      Set(ByVal value As Script)
        mobjScript = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpScript As Script, ByVal lpSuccess As Boolean, Optional ByVal lpDetails As String = "")
      MyBase.New(lpSuccess, lpDetails)
      Script = lpScript

      If AllMethodsSucceeded() Then
        Me.Details = String.Format("All {0} Methods succeeded", Script.Methods.Count)
        Me.Success = True
      Else
        ' This will need to be enhanced later
        Me.Details = "One or more methods failed"
        Me.Success = False
      End If
    End Sub

#End Region

#Region "Private Methods"

    Private Function AllMethodsSucceeded() As Boolean
      Try

        For Each lobjMethod As MethodBase In Script.Methods
          'If lobjMethod.LastResult.Success = False Then
          '  Return False
          'End If
        Next

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

  Public Class InvalidScriptException
    Inherits InvalidOperationException

#Region "Class Variables"

    Private mstrAttemptedScript As String = ""

#End Region

#Region "Public Properties"

    Public ReadOnly Property AttemptedScript() As String
      Get
        Return mstrAttemptedScript
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal message As String, ByVal lpAttemptedScript As String)
      Me.New(message, lpAttemptedScript, Nothing)
    End Sub

    Public Sub New(ByVal message As String, ByVal lpAttemptedScript As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrAttemptedScript = lpAttemptedScript
    End Sub

#End Region

  End Class

End Namespace