'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


Imports Documents.Core

#End Region

Namespace Scripting

  Public Class NowMethod
    Inherits MethodBase

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpKeyword As String)
      MyBase.Parameters.Add(PropertyFactory.Create(PropertyType.ecmString, "Keyword", lpKeyword))
    End Sub

#End Region

#Region "Public Methods"

    'Public Overrides Function Execute() As MethodResult

    '  Dim lobjReturnValue As Object = String.Empty

    '  Try

    '    lobjReturnValue = Parameters("Keyword").Value

    '    lobjReturnValue = ScriptMethods.Now(Parameters("Keyword").Value)

    '    Return New MethodResult(Me, True, lobjReturnValue, "NowMethod Succeeded")

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Return New MethodResult(Me, False, lobjReturnValue, "NowMethod Failed")
    '  End Try

    'End Function

    'Public Overrides Function LastResult() As MethodResult
    '  Try
    '    Return mobjLastResult
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  Finally

    '  End Try
    'End Function

#End Region

  End Class

End Namespace