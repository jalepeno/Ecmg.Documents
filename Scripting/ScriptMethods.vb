'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Scripting

  ''' <summary>
  ''' Contains shared methods for scripting within Cts applications
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ScriptMethods

#Region "Public Methods"

    '''' <summary>
    '''' Evaluates whether value starts with &lt;ctscript&gt; and ends with &lt;/ctscript&gt;
    '''' </summary>
    '''' <param name="lpValue"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function IsCtScript(ByVal lpValue As Object) As Boolean

    '  Try
    '    If lpValue.ToString.StartsWith("<ctscript>") AndAlso lpValue.ToString.EndsWith("</ctscript>") Then
    '      Return True
    '    End If

    '    Return False

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Function

    '''' <summary>
    '''' Returns the value inside the ctscript block
    '''' </summary>
    '''' <param name="lpScriptBlock"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function GetCtScriptValue(ByVal lpScriptBlock As String) As String

    '  Try
    '    Dim lstrScriptValue As String
    '    lstrScriptValue = lpScriptBlock.Replace("<ctscript>", "")
    '    lstrScriptValue = lstrScriptValue.Replace("</ctscript>", "")
    '    Return lstrScriptValue
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '  End Try

    '  Return lpScriptBlock

    'End Function

    'Public Shared Function GetValueFromScript(ByVal lpScript As String) As Object
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ''' <summary>
    ''' Takes variations of the Now keyword as a string and returns the current value
    ''' </summary>
    ''' <param name="lpNowKeyword">A case insensitive string variation of the Now keyword
    ''' 
    ''' Possible Values are
    ''' Now
    ''' Now.ToUniversalTime, Now.ToUtc, Now.Utc
    ''' Now.Date
    ''' Today
    ''' Now.ToUniversalTime.Date, Now.ToUtc.Date, Now.Utc.Date
    ''' Now.Day
    ''' Now.Minute
    ''' Now.Month
    ''' Now.Millisecond
    ''' Now.Hour
    ''' Now.Second
    ''' Now.DayOfYear
    ''' Now.Ticks
    ''' Now.DayOfWeek, Now.DayOfWeek.ToString
    ''' Now.ToLongDateString
    ''' Now.ToLongTimeString
    ''' Now.ToShortDateString
    ''' Now.ToShortTimeString
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Now(ByVal lpNowKeyword As String) As Object
      Try

        Dim lobjReturnValue As Object


        Select Case lpNowKeyword.ToLower
          Case "now"
            lobjReturnValue = Date.Now
          Case "now.touniversaltime", "now.toutc", "now.utc"
            lobjReturnValue = Date.Now.ToUniversalTime
          Case "now.date"
            lobjReturnValue = Date.Now.Date
          Case "today", "now.today"
            lobjReturnValue = Date.Today
          Case "now.touniversaltime.date", "now.toutc.date", "now.utc.date"
            lobjReturnValue = Date.Now.ToUniversalTime.Date
          Case "now.day"
            lobjReturnValue = Date.Now.Day
          Case "now.minute"
            lobjReturnValue = Date.Now.Minute
          Case "now.month"
            lobjReturnValue = Date.Now.Month
          Case "now.year"
            lobjReturnValue = Date.Now.Year
          Case "now.millisecond"
            lobjReturnValue = Date.Now.Millisecond
          Case "now.hour"
            lobjReturnValue = Date.Now.Hour
          Case "now.second"
            lobjReturnValue = Date.Now.Second
          Case "now.dayofyear"
            lobjReturnValue = Date.Now.DayOfYear
          Case "now.ticks"
            lobjReturnValue = Date.Now.Ticks
          Case "now.dayofweek"
            lobjReturnValue = Date.Now.DayOfWeek
          Case "now.dayofweek.tostring"
            lobjReturnValue = Date.Now.DayOfWeek.ToString
          Case "now.tolongdatestring"
            lobjReturnValue = Date.Now.ToLongDateString
          Case "now.tolongtimestring"
            lobjReturnValue = Date.Now.ToLongTimeString
          Case "now.toshortdatestring"
            lobjReturnValue = Date.Now.ToShortDateString
          Case "now.toshorttimestring"
            lobjReturnValue = Date.Now.ToShortTimeString
          Case Else
            lobjReturnValue = lpNowKeyword
        End Select

        Return lobjReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Depending on the target property type, will try to resolve the date keyword to the current value
    ''' </summary>
    ''' <param name="lpTargetPropertyType">The destination property type 
    ''' (using a valid Ecmg.Cts.Core.PropertyType enumeration value)</param>
    ''' <param name="lpNowKeyword">A case insensitive string variation of the Now keyword
    ''' 
    ''' Possible Values are
    ''' Now
    ''' Now.ToUniversalTime, Now.ToUtc, Now.Utc
    ''' Now.Date
    ''' Today
    ''' Now.ToUniversalTime.Date, Now.ToUtc.Date, Now.Utc.Date
    ''' Now.Day
    ''' Now.Minute
    ''' Now.Month
    ''' Now.Millisecond
    ''' Now.Hour
    ''' Now.Second
    ''' Now.DayOfYear
    ''' Now.Ticks
    ''' Now.DayOfWeek, Now.DayOfWeek.ToString
    ''' Now.ToLongDateString
    ''' Now.ToLongTimeString
    ''' Now.ToShortDateString
    ''' Now.ToShortTimeString
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Now(ByVal lpTargetPropertyType As Integer, ByVal lpNowKeyword As String) As Object
      Try

        Dim lobjReturnValue As Object = lpNowKeyword
        Dim lenuPropertyType As Core.PropertyType = lpTargetPropertyType

        Dim lstrNowKeyword As String = lpNowKeyword.Replace("()", String.Empty)

        Select Case lenuPropertyType
          Case Core.PropertyType.ecmDate
            If IsDate(lstrNowKeyword) = False Then
              ' If the property type is date and the value is a keyword 
              '   based value let's try to translate it here
              ' This may be another way of providing the value
              lobjReturnValue = Now(lstrNowKeyword)
            End If
          Case Core.PropertyType.ecmLong, Core.PropertyType.ecmDouble
            If IsNumeric(lstrNowKeyword) = False Then
              lobjReturnValue = Now(lstrNowKeyword)
            End If
          Case Core.PropertyType.ecmString
            If lstrNowKeyword.ToLower.StartsWith("now") _
              OrElse lstrNowKeyword.ToLower = "today" Then
              lobjReturnValue = Now(lstrNowKeyword)
            End If
        End Select

        Return lobjReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function


#End Region

  End Class

End Namespace
