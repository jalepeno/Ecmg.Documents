'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations
  ''' <summary>
  ''' Parses the source property and retrieves a part of the value to apply to the
  ''' destination property.
  ''' </summary>
  <Serializable()>
  Public Class DataParser
    Inherits DataLookup
    Implements IPropertyLookup

#Region "Class Constants"

    Friend Const PARAM_GENERATE_CDF As String = "GenerateCDF"
    Friend Const PARAM_PARSE_STYLE As String = "ParseStyle"
    Friend Const PARAM_PARSE_ARGUMENTS As String = "ParseStyleArguments"

#End Region

#Region "Public Enumerations"

    Public Enum PartEnum
      <Description("Takes the value of the source document property exactly as it exists in the current document and copies it into the destination property.")>
      Complete
      <Description("Evaluates the value of the source document property as it exists in the current document against the regular expression and copies the matches into the destination property. If the destination is a multi-valued property then each match will be copied into a separate value.")>
      Regex
      <Description("Takes the value of the source document property in the current document and concatenates it with the destination property. A delimiter may be specified to place between the two values if desired.")>
      Concatenate
      <Description("Concatenates the source property to the destination property only if the source property currently has a certain value. It uses the specified delimiter if provided.")>
      ConditionalConcatenate
      <Description("Concatenates the literal value to the destination property. It uses the specified delimiter if provided.")>
      ConcatenateLiteral
      <Description("Concatenates the literal value to the destination property only if the source property currently has a certain value. It uses the specified delimiter if provided. ")>
      ConditionalConcatenateLiteral
      <Description("Gets the value of the source property only if the destination property currently has a certain value and copies it into the destination property.")>
      Conditional
      <Description("Undefined")>
      ConditionalLiteral
      <Description("Takes the specified number of characters from the left of the value of the source document property in the current document and copies it into the destination property.")>
      Left
      <Description("Takes the specified number of characters from the right of the value of the source document property in the current document and copies it into the destination property.")>
      Right
      <Description("Takes the characters starting from the from the right of the value of the source document property in the current document and copies it into the destination property.")>
      SubString
      <Description("Removes the specified number of characters from the value of the source document property in the current document and copies the result into the destination property.")>
      RemoveLeft
      <Description("Removes the specified number of characters from the value of the source document property in the current document and copies the result into the destination property.")>
      RemoveRight
      <Description("Takes the value of the source document property in the current document, removes any leading or trailing spaces and copies the result into the destination property.")>
      Trim
      <Description("Takes the value of the source document property in the current document, removes any leading spaces and copies the result into the destination property.")>
      TrimLeft
      <Description("Takes the value of the source document property in the current document, removes any trailing spaces and copies the result into the destination property.")>
      TrimRight
      <Description("Takes the value of the source document property up to the last instance of the specified character in the current document and copies it into the destination property.")>
      RemoveAfterLastInstanceOfCharacter
      <Description("Takes the value of the source document property after the last instance of the specified character in the current document and copies it into the destination property.")>
      RemoveAllBeforeLastInstanceOfCharacter
      <Description("Takes the value of the source document property exactly as it exists in the current document and copies it into the destination property.")>
      Replace
      <Description("Takes the value of the source document property, makes all of the characters upper case and copies the result into the destination property.")>
      Upper
      <Description("Takes the value of the source document property, makes all of the characters lower case and copies the result into the destination property.")>
      Lower
      <Description("Takes the value of the source document property, makes the characters lower or upper case as custom dictates and copies the result into the destination property.")>
      Proper
      <Description("Takes one value from multi-valued source property and copies it into the destination property.")>
      Index
      <Description("Puts the Source Property in front of the Destination Property.")>
      Prefix
    End Enum

#End Region

#Region "Class Variables"

    Private mstrPart As String
    Private mobjSourceProperty As LookupProperty
    Private mobjDestinationProperty As LookupProperty

#End Region

#Region "Public Properties"

    ''' <summary>The property containing the source value.</summary>
    Public Property SourceProperty() As LookupProperty Implements IPropertyLookup.SourceProperty
      Get
        Return mobjSourceProperty
      End Get
      Set(ByVal Value As LookupProperty)
        mobjSourceProperty = Value
      End Set
    End Property

    ''' <summary>The property to update.</summary>
    Public Property DestinationProperty() As LookupProperty Implements IPropertyLookup.DestinationProperty
      Get
        Return mobjDestinationProperty
      End Get
      Set(ByVal Value As LookupProperty)
        mobjDestinationProperty = Value
      End Set
    End Property

    ''' <example>
    '''	<para><font color="#A31515" size="2">COMPLETE</font></para>
    '''	<para><font color="green" size="2">Get the complete value.</font></para>
    '''	<para><font color="#A31515" size="2">CONCATENATE</font></para>
    '''	<para><font color="green" size="2">Concatenate the source property to the
    '''    destination property,</font>
    '''		<font color="green" size="2">use the specified
    '''    delimiter if provided.</font></para>
    '''	<para><font color="#A31515" size="2">CONDITIONAL CONCATENATE</font></para>
    '''	<para><font color="green" size="2">Concatenate the source property to the
    '''    destination property only</font>
    '''		<font color="green" size="2">if the source
    '''    property currently has a certain value,</font>
    '''		<font color="green" size="2">use the
    '''    specified delimiter if provided.</font></para>
    '''	<para><font color="green" size="2">Expected 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:=</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:NOT EQUALS</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:STARTS WITH</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:ENDS WITH</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:CONTAINS</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:NOT EMPTY</font></para>
    '''	<para><font color="green" size="2">Optional 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:DESTINATION EMPTY</font></para>
    '''	<para><font color="#A31515" size="2">CONDITIONAL CONCATENATE LITERAL</font></para>
    '''	<para><font color="green" size="2">Concatenate the literal value to the destination
    '''    property only</font>
    '''		<font color="green" size="2">if the source property currently
    '''    has a certain value.</font>
    '''		<font color="green" size="2">Use the specified
    '''    delimiter if provided.</font></para>
    '''	<para><font color="green" size="2">Expected 'CONDITIONAL
    '''    CONCATENATE:Delmiter:Conditional Source Value:Literal Value</font></para>
    '''	<para><font color="#A31515" size="2">CONDITIONAL</font></para>
    '''	<para><font color="green" size="2">Get the source property only if the destination
    '''    property currently has a certain value.</font></para>
    '''	<para><font color="#A31515" size="2">CONDITIONAL LITERAL</font></para>
    '''	<para><font color="green" size="2">Apply a literal value only if the destination
    '''    property currently has a certain value.</font></para>
    '''	<para><font color="#A31515" size="2">LEFT</font></para>
    '''	<para><font color="green" size="2">Get just the part we want from the left of the
    '''    value.</font></para>
    '''	<para><font color="#A31515" size="2">RIGHT</font></para>
    '''	<para><font color="green" size="2">Get just the part we want from the right of the
    '''    value.</font></para>
    '''	<para><font color="#A31515" size="2">MID</font></para>
    '''	<para><font color="green" size="2">Get just the part we want from the middle of the
    '''    value.</font></para>
    '''	<para><font color="#A31515" size="2">REMOVE LEFT</font></para>
    '''	<para><font color="green" size="2">Remove the number of characters we don't want
    '''    from the left of the value.</font></para>
    '''	<para><font color="#A31515" size="2">REMOVE RIGHT</font></para>
    '''	<para><font color="green" size="2">Remove the number of characters we don't want
    '''    from the right of the value.</font></para>
    '''	<para><font color="#A31515" size="2">TRIM LEFT</font></para>
    '''	<para><font color="green" size="2">Remove the part we don't want from the left of
    '''    the value.</font></para>
    '''	<para><font color="#A31515" size="2">TRIM RIGHT</font></para>
    '''	<para><font color="green" size="2">Remove the part we don't want from the right of
    '''    the value.</font></para>
    '''	<para><font color="#A31515" size="2">REMOVE AFTER LAST INSTANCE OF
    '''    CHARACTER.</font></para>
    '''	<para><font color="green" size="2">Remove everything after the last instance of the
    '''    specified character.</font>
    '''		<font color="green" size="2">Good for removing file
    '''    extensions.</font></para>
    '''	<para><font color="#A31515" size="2">REPLACE</font></para>
    '''	<para><font color="green" size="2">Replace specified characters with supplied
    '''    characters.</font></para>
    '''	<para><font color="#A31515" size="2">UPPER</font></para>
    '''	<para><font color="green" size="2">Make all characters upper case.</font></para>
    '''	<para><font color="#A31515" size="2">LOWER</font></para>
    '''	<para><font color="green" size="2">Make all characters lower case.</font></para>
    ''' </example>
    ''' <summary>The part of the value to retrieve.</summary>
    Public Property Part() As String
      Get
        Return mstrPart
      End Get
      Set(ByVal Value As String)
        mstrPart = Value
      End Set
    End Property

#End Region

#Region "Constructors"

    ''' <summary>Default Constructor</summary>
    Public Sub New()
      MyBase.New(LookupType.Parser)
    End Sub

    Public Sub New(ByVal lpSourceProperty As LookupProperty,
                   ByVal lpDestinationProperty As LookupProperty,
                   ByVal lpPart As String)

      MyBase.New(LookupType.Parser)
      SourceProperty = lpSourceProperty
      DestinationProperty = lpDestinationProperty
      Part = lpPart

    End Sub

#End Region

#Region "Public Methods"

    Friend Shared Function GetPartStyle(lpPartString As String, ByRef lpArguments As String) As PartEnum
      Try

        Dim lobjPropertyPartArguments() As Object
        Dim lstrOperation As String = String.Empty
        Dim lenuPartEnum As PartEnum

        ' Parse out the operation from the 'Part' property
        lobjPropertyPartArguments = Split(lpPartString, ":")
        lstrOperation = lobjPropertyPartArguments(0).ToString.ToUpper

        If lobjPropertyPartArguments.Length > 1 Then
          lpArguments = lpPartString.Substring(lstrOperation.Length + 1)
        End If

        Select Case lstrOperation
          Case "COMPLETE"
            lenuPartEnum = PartEnum.Complete
          Case "REGEX"
            lenuPartEnum = PartEnum.Regex
          Case "CONCATENATE"
            lenuPartEnum = PartEnum.Concatenate
          Case "CONDITIONAL CONCATENATE"
            lenuPartEnum = PartEnum.ConditionalConcatenate
          Case "CONCATENATE LITERAL"
            lenuPartEnum = PartEnum.ConcatenateLiteral
          Case "CONDITIONAL CONCATENATE LITERAL"
            lenuPartEnum = PartEnum.ConditionalConcatenateLiteral
          Case "CONDITIONAL"
            lenuPartEnum = PartEnum.Conditional
          Case "CONDITIONAL LITERAL"
            lenuPartEnum = PartEnum.ConditionalLiteral
          Case "LEFT"
            lenuPartEnum = PartEnum.Left
          Case "RIGHT"
            lenuPartEnum = PartEnum.Right
          Case "MID", "SUBSTRING"
            lenuPartEnum = PartEnum.SubString
          Case "REMOVE LEFT"
            lenuPartEnum = PartEnum.RemoveLeft
          Case "REMOVE RIGHT"
            lenuPartEnum = PartEnum.RemoveRight
          Case "TRIM"
            lenuPartEnum = PartEnum.Trim
          Case "TRIM LEFT"
            lenuPartEnum = PartEnum.TrimLeft
          Case "TRIM RIGHT"
            lenuPartEnum = PartEnum.TrimRight
          Case "REMOVE AFTER LAST INSTANCE OF CHARACTER"
            lenuPartEnum = PartEnum.RemoveAfterLastInstanceOfCharacter
          Case "REMOVE ALL BEFORE LAST INSTANCE OF CHARACTER"
            lenuPartEnum = PartEnum.RemoveAllBeforeLastInstanceOfCharacter
          Case "REPLACE"
            lenuPartEnum = PartEnum.Replace
          Case "UPPER"
            lenuPartEnum = PartEnum.Upper
          Case "LOWER"
            lenuPartEnum = PartEnum.Lower
          Case "TITLE", "PROPER"
            lenuPartEnum = PartEnum.Proper
          Case "INDEX"
            lenuPartEnum = PartEnum.Index
          Case "PREFIX"
            lenuPartEnum = PartEnum.Prefix
          Case Else
            lenuPartEnum = PartEnum.Complete

        End Select

        Return lenuPartEnum

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Functions"

    ''' <summary>
    ''' Gets a value if the specified condition is met
    ''' </summary>
    ''' <param name="lpMetaHolder">The object from which to retrieve the source property</param>
    ''' <param name="lpPropertyPartArguments">If it is a literal then it the expected PropertyPartArguments are
    ''' CONDITIONAL LITERAL:lstrCondition:lstrLiteral:false
    ''' If it is not a literal then the expected PropertyPartArguments are
    ''' CONDITIONAL:lstrCondition:false
    ''' </param>
    ''' <param name="lpLiteral"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetConditionalValue(ByVal lpMetaHolder As Core.IMetaHolder,
                                     ByVal lpPropertyPartArguments() As Object,
                                     Optional ByVal lpLiteral As Boolean = False) As String

      ' If it is a literal then it the expected PropertyPartArguments are
      '   CONDITIONAL LITERAL:lstrCondition:lstrLiteral:false
      ' If it is not a literal then the expected PropertyPartArguments are
      '   CONDITIONAL:lstrCondition:false

      ' Get the source property only of the destination property currently has a certain value
      Dim lstrDestinationPropertyValue As String = GetProperty(DestinationProperty, lpMetaHolder).Value
      Dim lstrCondition As String
      Dim lstrSourcePropertyValue As String
      Dim lstrLiteralValue As String = ""
      Dim lblnCaseSensitive As Boolean = False

      lstrSourcePropertyValue = GetProperty(SourceProperty, lpMetaHolder).Value

      ' Get the conditional value
      Try
        lstrCondition = lpPropertyPartArguments(1)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' We were not able to get the condition
        ' Do not get the conditional value
        Return lstrDestinationPropertyValue
      End Try

      ' Get the literal value if needed
      If lpLiteral = True Then
        ' It is literal

        ' The third part is the literal value
        If lpPropertyPartArguments.Length > 2 Then
          lstrLiteralValue = CType(lpPropertyPartArguments(2), String)
        Else
          ' The literal is not specified
          ' Do not get the conditional value
          Return lstrDestinationPropertyValue
        End If

        ' Is the condition case sensitive?
        ' The fourth optional part is 'true/false' for case sensitivity
        If lpPropertyPartArguments.Length > 3 Then
          Try
            lblnCaseSensitive = CType(lpPropertyPartArguments(3), Boolean)
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            lblnCaseSensitive = False
          End Try
        Else
          lblnCaseSensitive = False
        End If

      Else
        ' It is not literal

        ' Is the condition case sensitive?
        ' The third optional part is 'true/false' for case sensitivity
        If lpPropertyPartArguments.Length > 2 Then
          Try
            lblnCaseSensitive = CType(lpPropertyPartArguments(2), Boolean)
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            lblnCaseSensitive = False
          End Try
        Else
          lblnCaseSensitive = False
        End If
      End If


      If lblnCaseSensitive = True Then
        If lstrDestinationPropertyValue = lpPropertyPartArguments(1).ToString Then
          If lpLiteral = False Then
            Return lstrSourcePropertyValue
          Else
            Return lstrLiteralValue
          End If
        Else
          Return lstrDestinationPropertyValue
        End If
      Else
        If lstrDestinationPropertyValue.ToUpper = lpPropertyPartArguments(1).ToString.ToUpper Then
          If lpLiteral = False Then
            Return lstrSourcePropertyValue
          Else
            Return lstrLiteralValue
          End If
        Else
          Return lstrDestinationPropertyValue
        End If
      End If

    End Function

    ''' <summary>
    ''' Gets all of the applicable matches given the specified document and regular expression
    ''' </summary>
    ''' <param name="lpSourcePropertyValue">The value to test against</param>
    ''' <returns></returns>
    ''' <remarks>Expected Part pattern 'REGEX:BEGIN*Regular Expression*END:BEGIN_REPLACE*Replace Expression*END_REPLACE:Return Index</remarks>
    Private Function GetRegexValue(ByVal lpSourcePropertyValue As String) As String
      Try

        Dim lintReturnIndex As Integer
        'Dim lobjMatchCollection As MatchCollection
        Dim lobjMatches As List(Of String)

        'lobjMatchCollection = GetRegexMatches(lpSourcePropertyValue)
        lobjMatches = GetRegexMatches(lpSourcePropertyValue)

        'If lobjMatchCollection Is Nothing Then
        '  Return String.Empty
        'End If

        If lobjMatches.Count = 0 Then
          'Return String.Empty
          'If no matches found then return original value
          Return lpSourcePropertyValue
        End If

        'If lpPropertyPartArguments.Length > 2 Then
        '  If IsNumeric(lpPropertyPartArguments(2)) = False Then
        '    Throw New ArgumentException(String.Format("The value '{0}' supplied for return index is not numeric.  Please supply a non-negative integer.", lpPropertyPartArguments(2)))
        '  Else
        '    lintReturnIndex = lpPropertyPartArguments(2)
        '  End If
        'End If

        'If (lobjMatchCollection.Count - 1) < lintReturnIndex Then
        '  ' Write a warning in the log and ignore the first return index
        '  Dim lstrMessage As String = String.Format("The value of '{0}' supplied for the return index is greater than the regular expression match count of '{1}' for the regular expression of '{2}' used to test the value of '{3}'.  A value of zero for the first match value will be assumed.", _
        '              lintReturnIndex, lobjMatchCollection.Count, lstrRegularExpression, lpSourcePropertyValue)
        '  ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Warning, 1001)
        '  lintReturnIndex = 0
        'End If

        'For lintIndexCounter As Integer = lintFirstReturnIndex To lobjMatchCollection.Count - 1
        '  lobjReturnArray.Add(lobjMatchCollection(lintIndexCounter).Value)
        'Next

        'Return lobjReturnArray

        'Return lobjMatchCollection(lintReturnIndex).Value
        Return lobjMatches(lintReturnIndex)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Return an empty string
        Return String.Empty
      End Try
    End Function

    Private Function GetRegexValues(ByVal lpSourcePropertyValue As String) As Values
      Try

        'Dim lobjMatchCollection As MatchCollection
        Dim lobjMatches As List(Of String)
        Dim lobjValues As New Values

        'lobjMatchCollection = GetRegexMatches(lpSourcePropertyValue)
        lobjMatches = GetRegexMatches(lpSourcePropertyValue)

        'If lobjMatchCollection Is Nothing Then
        '  Return New Values
        'End If

        If lobjMatches.Count = 0 Then
          Return New Values
        End If

        'For Each lobjMatch As Match In lobjMatchCollection
        '  If lobjValues.Contains(lobjMatch.Value) = False Then
        '    lobjValues.Add(lobjMatch.Value)
        '  End If
        'Next

        For Each lstrMatch As String In lobjMatches
          If lobjValues.Contains(lstrMatch) = False Then
            lobjValues.Add(lstrMatch)
          End If
        Next

        Return lobjValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Return an empty Values collection
        Return New Values
      End Try
    End Function

    ''' <summary>
    ''' Gets all of the applicable matches given the specified document and regular expression
    ''' </summary>
    ''' <param name="lpSourcePropertyValue">The value to test against</param>
    ''' <returns></returns>
    ''' <remarks>Expected Part pattern 'REGEX:BEGIN*Regular Expression*END:BEGIN_REPLACE*Replace Expression*END_REPLACE:Return Index</remarks>
    Private Function GetRegexMatches(ByVal lpSourcePropertyValue As String) As List(Of String)
      Try

        'Dim lintReturnIndex As Integer
        Dim lobjRegex As Regex
        Dim lobjMatchCollection As MatchCollection
        'Dim lobjReturnArray As New ArrayList
        Dim lstrReplacementString As String
        Dim lstrReplacedValue As String
        Dim lobjReturnList As New List(Of String)

        ' Make sure we got the expected arguments 
        If Part Is Nothing OrElse Part.Length = 0 Then
          Throw New InvalidOperationException("No value available for PART in GetRegexValues")
        End If

        Dim lstrPropertyPartArguments() As String = Split(Part, ":")

        'lobjRegex = New Regex("^regex", _
        '          RegexOptions.IgnoreCase Or _
        '          Text.RegularExpressions.RegexOptions.CultureInvariant Or _
        '          Text.RegularExpressions.RegexOptions.Compiled)

        If Part.ToUpper.StartsWith("REGEX") = False Then
          Throw New ArgumentException(String.Format(
                  "The expected value for the first segment of the delimited PART string is 'REGEX', not '{0}",
                  lstrPropertyPartArguments(0)), "lpPropertyPartArguments")
        End If


        ' Get the regular expression
        lobjRegex = New Regex("(?<=regex:begin\*).*?(?=\*end)",
                  RegexOptions.IgnoreCase Or
                  Text.RegularExpressions.RegexOptions.CultureInvariant Or
                  Text.RegularExpressions.RegexOptions.Compiled)

        lobjMatchCollection = lobjRegex.Matches(Part)
        If lobjMatchCollection.Count < 1 Then
          Throw New ArgumentException("Must supply value for regular expression in delimited part string.  Expected 'REGEX:BEGIN*regular expression*END:optional BEGIN_REPLACE*Replace Expression*END_REPLACE:optional first return index'", "lpPropertyPartArguments")
        End If

        Dim lstrRegularExpression As String = lobjMatchCollection(0).Value


        lobjRegex = New Regex(lstrRegularExpression,
                  RegexOptions.IgnoreCase Or
                  Text.RegularExpressions.RegexOptions.CultureInvariant Or
                  Text.RegularExpressions.RegexOptions.Compiled)

        ' Capture all matches in the input text
        lobjMatchCollection = lobjRegex.Matches(lpSourcePropertyValue)

        If lobjMatchCollection.Count = 0 Then
          ' We did not find anything, just return the empty list
          Return lobjReturnList
        End If


        'If lpPropertyPartArguments.Length > 2 Then
        '  If IsNumeric(lpPropertyPartArguments(2)) = False Then
        '    Throw New ArgumentException(String.Format("The value '{0}' supplied for return index is not numeric.  Please supply a non-negative integer.", lpPropertyPartArguments(2)))
        '  Else
        '    lintReturnIndex = lpPropertyPartArguments(2)
        '  End If
        'End If

        'If (lobjMatchCollection.Count - 1) < lintReturnIndex Then
        '  ' Write a warning in the log and ignore the first return index
        '  Dim lstrMessage As String = String.Format("The value of '{0}' supplied for the return index is greater than the regular expression match count of '{1}' for the regular expression of '{2}' used to test the value of '{3}'.  A value of zero for the first match value will be assumed.", _
        '              lintReturnIndex, lobjMatchCollection.Count, lstrRegularExpression, lpSourcePropertyValue)
        '  ApplicationLogging.WriteLogEntry(lstrMessage, TraceEventType.Warning, 1001)
        '  lintReturnIndex = 0
        'End If

        'For lintIndexCounter As Integer = lintFirstReturnIndex To lobjMatchCollection.Count - 1
        '  lobjReturnArray.Add(lobjMatchCollection(lintIndexCounter).Value)
        'Next

        'Return lobjReturnArray

        ' See if a replacement string was specified
        Dim lobjRegexReplacement As New Regex("(?<=begin_replace\*).*?(?=\*end_replace)",
                  RegexOptions.IgnoreCase Or
                  Text.RegularExpressions.RegexOptions.CultureInvariant Or
                  Text.RegularExpressions.RegexOptions.Compiled)

        Dim lobjReplacementStringMatch As Match = lobjRegexReplacement.Match(Part)

        If lobjReplacementStringMatch.Length = 0 Then
          ' There was no replacement specifier

          ' Add the matches as they are to the return list
          For Each lobjMatch As Match In lobjMatchCollection
            lobjReturnList.Add(lobjMatch.Value)
          Next

        Else

          ' We found a replacement specifier
          lstrReplacementString = lobjReplacementStringMatch.Value

          ' See if we have a return specifier
          Dim lobjRegexReturnStyle As New Regex("(?<=\^return\().*?(?=\))",
                    RegexOptions.IgnoreCase Or
                    Text.RegularExpressions.RegexOptions.CultureInvariant Or
                    Text.RegularExpressions.RegexOptions.Compiled)

          Dim lobjReturnStyle As Match = lobjRegexReturnStyle.Match(Part)

          If lobjReturnStyle IsNot Nothing Then
            Select Case lobjReturnStyle.Value.ToUpper
              Case "MATCHONLY"
                ' Only return the matched portion
                lstrReplacedValue = Regex.Replace(lobjMatchCollection(0).Value, lstrRegularExpression, lstrReplacementString)
              Case "ALL"
                ' Return the entire value with the replacement
                lstrReplacedValue = Regex.Replace(lpSourcePropertyValue, lstrRegularExpression, lstrReplacementString)
              Case Else
                ' Return the entire value with the replacement
                lstrReplacedValue = Regex.Replace(lpSourcePropertyValue, lstrRegularExpression, lstrReplacementString)
            End Select

            lobjReturnList.Add(lstrReplacedValue)

          End If

        End If

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Return an empty string
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="lpSourceProperty"></param>
    ''' <param name="lpPropertyPartArguments"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndexedValue(ByVal lpSourceProperty As Core.ECMProperty, ByVal lpPropertyPartArguments() As Object) As Object
      Try

        If lpSourceProperty Is Nothing Then
          Throw New ArgumentNullException("lpSourceProperty")
        End If

        If lpPropertyPartArguments Is Nothing Then
          Throw New ArgumentNullException("lpPropertyPartArguments")
        End If

        ' Make sure we have an index value
        If lpPropertyPartArguments.Length < 2 Then
          Throw New ArgumentException("Must supply value for index in delimited part string.  Expected 'INDEX:Index Number'",
                                      "lpPropertyPartArguments")
        End If

        ' Make sure the index number is numeric
        If IsNumeric(lpPropertyPartArguments(1)) = False Then
          Throw New ArgumentException(String.Format(
            "Value for index must be numeric. '{0}' is not a valid index value. Expected 'INDEX:Index Number'",
            lpPropertyPartArguments(1), "lpPropertyPartArguments"))
        End If

        Dim lintIndexNumber As Integer = lpPropertyPartArguments(1)

        Return GetIndexedValue(lpSourceProperty, lintIndexNumber)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Return an empty string
        Return String.Empty
      End Try
    End Function

    ''' <summary>
    ''' Gets the indexed value from the source property
    ''' </summary>
    ''' <param name="lpSourceProperty">The source property to get the value from</param>
    ''' <param name="lpIndex">The index number for the property value to retrieve (one-based)</param>
    ''' <returns></returns>
    ''' <remarks>Intended for use with multi-valued source properties</remarks>
    Private Function GetIndexedValue(ByVal lpSourceProperty As Core.ECMProperty, ByVal lpIndex As Integer) As Object
      Try

        If lpSourceProperty Is Nothing Then
          Throw New ArgumentNullException("lpSourceProperty")
        End If

        If lpIndex < 1 Then
          Throw New ArgumentOutOfRangeException("lpIndex", lpIndex,
            String.Format("The value '{0}' specified for lpIndex is invalid.  Only values greater than one are allowed.", lpIndex))
        End If

        If lpSourceProperty.Cardinality <> Cardinality.ecmMultiValued Then
          Throw New InvalidOperationException(
            String.Format("Operation not valid on single-valued property '{0}', use only on multi-valued properties.",
                          lpSourceProperty.Name))
        End If

        If (lpSourceProperty.Values.Count < lpIndex) Then
          Throw New ArgumentOutOfRangeException("lpIndex",
            String.Format("The value '{0}' specified for lpIndex is greater than the number of values '{1}' present in property '{2}'",
            lpIndex, lpSourceProperty.Values.Count, lpSourceProperty.Name))
        End If

        Return lpSourceProperty.Values(lpIndex - 1)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Return an empty string
        Return String.Empty
      End Try
    End Function

    'Private Function GetConditionalValue(ByVal lpDocument As Ecmg.Cts.Core.Document, _
    '                                     ByVal lpPropertyPartArguments() As Object, _
    '                                     Optional ByVal lpLiteral As Boolean = False) As String

    '  ' If it is a literal then it the expected PropertyPartArguments are
    '  '   CONDITIONAL LITERAL:lstrCondition:lstrLiteral:false
    '  ' If it is not a literal then the expected PropertyPartArguments are
    '  '   CONDITIONAL:lstrCondition:false


    '  ' Get the source property only of the destination property currently has a certain value
    '  Dim lstrDestinationPropertyValue As String = GetProperty(DestinationProperty, lpDocument).Value
    '  Dim lstrCondition As String
    '  Dim lstrSourcePropertyValue As String
    '  Dim lstrLiteralValue As String = ""
    '  Dim lblnCaseSensitive As Boolean = False

    '  lstrSourcePropertyValue = GetProperty(SourceProperty, lpDocument).Value

    '  ' Get the conditional value
    '  Try
    '    lstrCondition = lpPropertyPartArguments(1)
    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' We were not able to get the condition
    '    ' Do not get the conditional value
    '    Return lstrDestinationPropertyValue
    '  End Try

    '  ' Get the literal value if needed
    '  If lpLiteral = True Then
    '    ' It is literal

    '    ' The third part is the literal value
    '    If lpPropertyPartArguments.Length > 2 Then
    '      lstrLiteralValue = CType(lpPropertyPartArguments(2), String)
    '    Else
    '      ' The literal is not specified
    '      ' Do not get the conditional value
    '      Return lstrDestinationPropertyValue
    '    End If

    '    ' Is the condition case sensitive?
    '    ' The fourth optional part is 'true/false' for case sensitivity
    '    If lpPropertyPartArguments.Length > 3 Then
    '      Try
    '        lblnCaseSensitive = CType(lpPropertyPartArguments(3), Boolean)
    '      Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        lblnCaseSensitive = False
    '      End Try
    '    Else
    '      lblnCaseSensitive = False
    '    End If

    '  Else
    '    ' It is not literal

    '    ' Is the condition case sensitive?
    '    ' The third optional part is 'true/false' for case sensitivity
    '    If lpPropertyPartArguments.Length > 2 Then
    '      Try
    '        lblnCaseSensitive = CType(lpPropertyPartArguments(2), Boolean)
    '      Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        lblnCaseSensitive = False
    '      End Try
    '    Else
    '      lblnCaseSensitive = False
    '    End If
    '  End If


    '  If lblnCaseSensitive = True Then
    '    If lstrDestinationPropertyValue = lpPropertyPartArguments(1).ToString Then
    '      If lpLiteral = False Then
    '        Return lstrSourcePropertyValue
    '      Else
    '        Return lstrLiteralValue
    '      End If
    '    Else
    '      Return lstrDestinationPropertyValue
    '    End If
    '  Else
    '    If lstrDestinationPropertyValue.ToUpper = lpPropertyPartArguments(1).ToString.ToUpper Then
    '      If lpLiteral = False Then
    '        Return lstrSourcePropertyValue
    '      Else
    '        Return lstrLiteralValue
    '      End If
    '    Else
    '      Return lstrDestinationPropertyValue
    '    End If
    '  End If

    'End Function

    'Private Function GetConditionalValue(ByVal lpVersion As Ecmg.Cts.Core.Version, _
    '                                     ByVal lpPropertyPartArguments() As Object, _
    '                                     Optional ByVal lpLiteral As Boolean = False) As String

    '  ' If it is a literal then it the expected PropertyPartArguments are
    '  '   CONDITIONAL LITERAL:lstrCondition:lstrLiteral:false
    '  ' If it is not a literal then the expected PropertyPartArguments are
    '  '   CONDITIONAL:lstrCondition:false


    '  ' Get the source property only of the destination property currently has a certain value
    '  Dim lstrDestinationPropertyValue As String = GetProperty(DestinationProperty, lpVersion).Value
    '  Dim lstrCondition As String
    '  Dim lstrSourcePropertyValue As String
    '  Dim lstrLiteralValue As String = ""
    '  Dim lblnCaseSensitive As Boolean = False

    '  lstrSourcePropertyValue = GetProperty(SourceProperty, lpVersion).Value

    '  ' Get the conditional value
    '  Try
    '    lstrCondition = lpPropertyPartArguments(1)
    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' We were not able to get the condition
    '    ' Do not get the conditional value
    '    Return lstrDestinationPropertyValue
    '  End Try

    '  ' Get the literal value if needed
    '  If lpLiteral = True Then
    '    ' It is literal

    '    ' The third part is the literal value
    '    If lpPropertyPartArguments.Length > 2 Then
    '      lstrLiteralValue = CType(lpPropertyPartArguments(2), String)
    '    Else
    '      ' The literal is not specified
    '      ' Do not get the conditional value
    '      Return lstrDestinationPropertyValue
    '    End If

    '    ' Is the condition case sensitive?
    '    ' The fourth optional part is 'true/false' for case sensitivity
    '    If lpPropertyPartArguments.Length > 3 Then
    '      Try
    '        lblnCaseSensitive = CType(lpPropertyPartArguments(3), Boolean)
    '      Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        lblnCaseSensitive = False
    '      End Try
    '    Else
    '      lblnCaseSensitive = False
    '    End If

    '  Else
    '    ' It is not literal

    '    ' Is the condition case sensitive?
    '    ' The third optional part is 'true/false' for case sensitivity
    '    If lpPropertyPartArguments.Length > 2 Then
    '      Try
    '        lblnCaseSensitive = CType(lpPropertyPartArguments(2), Boolean)
    '      Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        lblnCaseSensitive = False
    '      End Try
    '    Else
    '      lblnCaseSensitive = False
    '    End If
    '  End If


    '  If lblnCaseSensitive = True Then
    '    If lstrDestinationPropertyValue = lpPropertyPartArguments(1).ToString Then
    '      If lpLiteral = False Then
    '        Return lstrSourcePropertyValue
    '      Else
    '        Return lstrLiteralValue
    '      End If
    '    Else
    '      Return lstrDestinationPropertyValue
    '    End If
    '  Else
    '    If lstrDestinationPropertyValue.ToUpper = lpPropertyPartArguments(1).ToString.ToUpper Then
    '      If lpLiteral = False Then
    '        Return lstrSourcePropertyValue
    '      Else
    '        Return lstrLiteralValue
    '      End If
    '    Else
    '      Return lstrDestinationPropertyValue
    '    End If
    '  End If

    'End Function

    ''' <summary>
    ''' Gets the usable value for the data parsing operation based on the operation specified in the Part property
    ''' </summary>
    ''' <param name="lpMetaHolder">The object from which to retrieve the source property</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetUsableValue(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      ' BMK GetUsableValue
      Dim lobjSourceProperty As Core.ECMProperty = Nothing
      Dim lobjDestinationProperty As Core.ECMProperty = Nothing

      Dim lobjPropertyPartArguments() As Object
      'Dim lstrOperation As String = String.Empty
      Dim lenuOperation As PartEnum = GetPartStyle(Part, String.Empty)

      Dim lstrSourcePropertyValue As String

      Dim lstrErrorMessage As String = String.Empty

      Try

        ' Parse out the operation from the 'Part' property
        lobjPropertyPartArguments = Split(Part, ":")
        'lstrOperation = lobjPropertyPartArguments(0).ToString.ToUpper

        lobjSourceProperty = GetProperty(SourceProperty, lpMetaHolder)

        If lenuOperation = PartEnum.Regex Then
          Try
            lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
          Catch PropEx As PropertyDoesNotExistException
            If PropEx.RequestedPropertyName = "ContentPath" Then
              If PropEx.Message.EndsWith("Version object") Then
                ApplicationLogging.WriteLogEntry("Unable to get ContentPath from Version object, going to the primary content element.", TraceEventType.Warning, 6813)
                lobjDestinationProperty = GetProperty(DestinationProperty, CType(lpMetaHolder, Version).GetPrimaryContent)
              ElseIf PropEx.Message.EndsWith("Document Object") Then
                ApplicationLogging.WriteLogEntry("Unable to get ContentPath from Document object, going to the primary content element of the latest version.", TraceEventType.Warning, 6814)
                lobjDestinationProperty = GetProperty(DestinationProperty, CType(lpMetaHolder, Document).GetLatestVersion.GetPrimaryContent)
              End If
            Else
              ApplicationLogging.LogException(PropEx, Reflection.MethodBase.GetCurrentMethod)
              ' Re-throw the exception to the caller
              Throw
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End If

        ' Get the baseline value for lstrSourcePropertyValue 
        Select Case lobjSourceProperty.Cardinality
          Case Cardinality.ecmSingleValued
            ' Get the value
            If TypeOf lobjSourceProperty.Value Is Guid Then
              lstrSourcePropertyValue = DirectCast(lobjSourceProperty.Value, Guid).ToString
            Else
              lstrSourcePropertyValue = lobjSourceProperty.Value
            End If
          Case Cardinality.ecmMultiValued
            ' Just get the first value
            ' We may have to change this later
            If lobjSourceProperty.Values.Count > 0 Then
              If TypeOf lobjSourceProperty.Values(0) Is Guid Then
                lstrSourcePropertyValue = DirectCast(lobjSourceProperty.Values(0), Guid).ToString
              Else
                lstrSourcePropertyValue = lobjSourceProperty.Values(0)
              End If
            Else
              lstrSourcePropertyValue = String.Empty
            End If
          Case Else
            ' This should never happen, but just in case, treat it as a single valued property
            lstrSourcePropertyValue = lobjSourceProperty.Value
        End Select

        Select Case lenuOperation
          Case PartEnum.Complete
            ' Do Nothing
            Return lstrSourcePropertyValue
          Case PartEnum.Concatenate
            ' Concatenate the source property to the destination property
            ' Use the specified delimiter if provided
            Return GetConcatenatedValue(lpMetaHolder, lstrSourcePropertyValue, lobjPropertyPartArguments)
          Case PartEnum.ConditionalConcatenate
            ' Concatenate the source property to the destination property only 
            '   if the source property currently has a certain value
            ' Use the specified delimiter if provided
            ' Expected 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:=
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:EQUALS
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:<>
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EQUALS
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:STARTS WITH
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:ENDS WITH
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:CONTAINS
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EMPTY
            '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:DESTINATION EMPTY
            Return GetConditionalConcatenatedValue(lpMetaHolder, lstrSourcePropertyValue, lobjPropertyPartArguments)
            ' <Added by: Ernie at: 9/10/2013-10:39:27 AM on machine: ERNIE-THINK>
          Case PartEnum.ConcatenateLiteral
            Return GetLiteralConcatenatedValue(lpMetaHolder, lobjPropertyPartArguments)
            ' </Added by: Ernie at: 9/10/2013-10:39:27 AM on machine: ERNIE-THINK>
          Case PartEnum.ConditionalConcatenateLiteral
            ' Concatenate the literal value to the destination property only 
            '   if the source property currently has a certain value
            ' Use the specified delimiter if provided
            'Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
            Return GetConditionalConcatenatedLiteral(lpMetaHolder, lstrSourcePropertyValue, lobjPropertyPartArguments)
          Case PartEnum.Conditional
            ' Get the source property only if the destination property currently has a certain value
            Return GetConditionalValue(lpMetaHolder, lobjPropertyPartArguments)
          Case PartEnum.ConditionalLiteral
            ' Apply a literal value only if the destination property currently has a certain value
            Return GetConditionalValue(lpMetaHolder, lobjPropertyPartArguments, True)
          Case PartEnum.Left
            ' Get just the part we want from the left of the value
            Return Left(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
          Case PartEnum.Right
            ' Get just the part we want from the right of the value
            Return Right(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
          Case PartEnum.SubString
            ' Get just the part we want from the middle of the value
            If lobjPropertyPartArguments.Length > 2 Then
              Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
            Else
              Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
            End If
          Case PartEnum.RemoveLeft
            ' Remove the number of characters we don't want from the left of the value
            Dim lintLeftLength As Integer = CType(lobjPropertyPartArguments(1), Integer)
            ' Make sure that we are not asking to remove more characters than are present
            If lintLeftLength > lstrSourcePropertyValue.Length Then
              lstrErrorMessage = String.Format("Unable to change property value: The number of characters specified '{0}' to remove from '{1}' is greater than the length of the source property value to remove them from '{2}', returning original value of '{3}'",
                                    lintLeftLength, lobjSourceProperty.Name, lstrSourcePropertyValue.Length, lstrSourcePropertyValue)
              ApplicationLogging.WriteLogEntry(lstrErrorMessage, TraceEventType.Warning, 1111)
              Return lstrSourcePropertyValue
            Else
              Return lstrSourcePropertyValue.Remove(0, lintLeftLength)
            End If

          Case PartEnum.RemoveRight
            ' Remove the number of characters we don't want from the right of the value
            Dim lintRightLength As Integer = CType(lobjPropertyPartArguments(1), Integer)
            ' Make sure that we are not asking to remove more characters than are present
            If lintRightLength > lstrSourcePropertyValue.Length Then
              lstrErrorMessage = String.Format("Unable to change property value (REMOVE RIGHT): The number of characters specified '{0}' to remove from '{1}' is greater than the length of the source property value to remove them from '{2}', returning original value of '{3}'",
                                    lintRightLength, lobjSourceProperty.Name, lstrSourcePropertyValue.Length, lstrSourcePropertyValue)
              ApplicationLogging.WriteLogEntry(lstrErrorMessage, TraceEventType.Warning, 1111)
              Return lstrSourcePropertyValue
            Else
              Return lstrSourcePropertyValue.Remove((lstrSourcePropertyValue.Length - lintRightLength), lintRightLength)
            End If

          Case PartEnum.Trim
            ' Remove the part we don't want from the left of the value
            If lobjPropertyPartArguments.Length = 1 Then
              ' No value was supplied for the character(s) to trim, default to a single space
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Both, " ")
            Else
              ' Trim the specified character(s)
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Both, lobjPropertyPartArguments(1))
            End If
          Case PartEnum.TrimLeft
            ' Remove the part we don't want from the left of the value
            If lobjPropertyPartArguments.Length = 1 Then
              ' No value was supplied for the character(s) to trim, default to a single space
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Left, " ")
            Else
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Left, lobjPropertyPartArguments(1))
            End If
          Case PartEnum.TrimRight
            ' Remove the part we don't want from the right of the value
            If lobjPropertyPartArguments.Length = 1 Then
              ' No value was supplied for the character(s) to trim, default to a single space
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Right, " ")
            Else
              Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Right, lobjPropertyPartArguments(1))
            End If
          Case PartEnum.RemoveAfterLastInstanceOfCharacter
            ' Remove everything after the last instance of the specified character
            ' Good for removing file extensions
            Return lstrSourcePropertyValue.Remove(lstrSourcePropertyValue.LastIndexOf(lobjPropertyPartArguments(1)))
          Case PartEnum.RemoveAllBeforeLastInstanceOfCharacter
            ' Remove everything after the last instance of the specified character
            ' Good for removing file extensions
            Return lstrSourcePropertyValue.Substring(lstrSourcePropertyValue.LastIndexOf(lobjPropertyPartArguments(1)))
          Case PartEnum.Replace
            ' Replace specified characters with supplied characters
            Return lstrSourcePropertyValue.Replace(lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
          Case PartEnum.Upper
            ' Make all characters upper case
            Return lstrSourcePropertyValue.ToUpper
          Case PartEnum.Lower
            ' Make all characters lower case
            Return lstrSourcePropertyValue.ToLower
          Case PartEnum.Proper
            ' Make all characters lower case
            Return StrConv(lstrSourcePropertyValue, VbStrConv.ProperCase)
          Case PartEnum.Regex
            ' Expected 'REGEX:Regular Expression:First Return Index
            If lobjDestinationProperty.Cardinality = Cardinality.ecmSingleValued Then
              Return GetRegexValue(lstrSourcePropertyValue)
            Else
              Return GetRegexValues(lstrSourcePropertyValue)
            End If
          Case PartEnum.Index
            ' Expected 'INDEX:Index Number'
            ' Index number is one based (1 is the first value)
            Return GetIndexedValue(lobjSourceProperty, lobjPropertyPartArguments)
          Case PartEnum.Prefix
            ' Puts the Source Property in front of the Destination Property
            ' Use the specified delimiter if provided
            Return GetPrefixedValue(lpMetaHolder, lstrSourcePropertyValue, lobjPropertyPartArguments)

          Case Else
            ' Do Nothing
            Return ""
        End Select

      Catch ex As Exception
        lstrErrorMessage = String.Format("Unable to get usable value for operation '{0}': {1}",
                        Part, ex.Message)
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 1999)
        Throw New InvalidOperationException(lstrErrorMessage, ex)
      End Try

    End Function

    'Private Function GetUsableValue(ByVal lpVersion As Ecmg.Cts.Core.Version) As String

    '  Dim lobjSourceProperty As Core.ECMProperty

    '  Dim lobjPropertyPartArguments() As Object
    '  Dim lstrOperation As String
    '  Dim lstrSourcePropertyValue As String

    '  lobjSourceProperty = GetProperty(SourceProperty, lpVersion)
    '  lstrSourcePropertyValue = lobjSourceProperty.Value

    '  Try
    '    lobjPropertyPartArguments = Split(Part, ":")
    '    lstrOperation = lobjPropertyPartArguments(0).ToString.ToUpper
    '    Select Case lstrOperation
    '      Case "COMPLETE"
    '        ' Do Nothing
    '        Return lstrSourcePropertyValue
    '      Case "CONCATENATE"
    '        ' Concatenate the source property to the destination property
    '        ' Use the specified delimiter if provided
    '        Return GetConcatenatedValue(lpVersion, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL CONCATENATE"
    '        ' Concatenate the source property to the destination property only 
    '        '   if the source property currently has a certain value
    '        ' Use the specified delimiter if provided
    '        ' Expected 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:=
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:EQUALS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:<>
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EQUALS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:STARTS WITH
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:ENDS WITH
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:CONTAINS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EMPTY
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:DESTINATION EMPTY
    '        Return GetConditionalConcatenatedValue(lpVersion, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL CONCATENATE LITERAL"
    '        ' Concatenate the literal value to the destination property only 
    '        '   if the source property currently has a certain value
    '        ' Use the specified delimiter if provided
    '        'Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
    '        Return GetConditionalConcatenatedLiteral(lpVersion, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL"
    '        ' Get the source property only if the destination property currently has a certain value
    '        Return GetConditionalValue(lpVersion, lobjPropertyPartArguments)
    '      Case "CONDITIONAL LITERAL"
    '        ' Apply a literal value only if the destination property currently has a certain value
    '        Return GetConditionalValue(lpVersion, lobjPropertyPartArguments, True)
    '      Case "LEFT"
    '        ' Get just the part we want from the left of the value
    '        Return Left(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '      Case "RIGHT"
    '        ' Get just the part we want from the right of the value
    '        Return Right(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '      Case "MID", "SUBSTRING"
    '        ' Get just the part we want from the middle of the value
    '        If lobjPropertyPartArguments.Length > 2 Then
    '          Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
    '        Else
    '          Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '        End If
    '      Case "REMOVE LEFT"
    '        ' Remove the number of characters we don't want from the left of the value
    '        Return lstrSourcePropertyValue.Remove(0, lobjPropertyPartArguments(1))
    '      Case "REMOVE RIGHT"
    '        ' Remove the number of characters we don't want from the right of the value
    '        Dim lintRightLength As Integer = CType(lobjPropertyPartArguments(1), Integer)
    '        Return lstrSourcePropertyValue.Remove((lstrSourcePropertyValue.Length - lintRightLength), lintRightLength)
    '      Case "TRIM"
    '        ' Remove the part we don't want from the left of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Both, lobjPropertyPartArguments(1))
    '      Case "TRIM LEFT"
    '        ' Remove the part we don't want from the left of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Left, lobjPropertyPartArguments(1))
    '      Case "TRIM RIGHT"
    '        ' Remove the part we don't want from the right of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Right, lobjPropertyPartArguments(1))
    '      Case "REMOVE AFTER LAST INSTANCE OF CHARACTER"
    '        ' Remove everything after the last instance of the specified character
    '        ' Good for removing file extensions
    '        Return lstrSourcePropertyValue.Remove(lstrSourcePropertyValue.LastIndexOf(lobjPropertyPartArguments(1)))
    '      Case "REPLACE"
    '        ' Replace specified characters with supplied characters
    '        Return lstrSourcePropertyValue.Replace(lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
    '      Case "UPPER"
    '        ' Make all characters upper case
    '        Return lstrSourcePropertyValue.ToUpper
    '      Case "LOWER"
    '        ' Make all characters lower case
    '        Return lstrSourcePropertyValue.ToLower
    '      Case Else
    '        ' Do Nothing
    '        Return ""
    '    End Select

    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Throw New InvalidOperationException("Unable to get usable value [" & ex.Message & "]", ex)
    '  End Try

    'End Function

    'Private Function GetUsableValue(ByVal lpDocument As Ecmg.Cts.Core.Document) As String

    '  Dim lobjSourceProperty As Core.ECMProperty

    '  Dim lobjPropertyPartArguments() As Object
    '  Dim lstrOperation As String
    '  Dim lstrSourcePropertyValue As String

    '  lobjSourceProperty = GetProperty(SourceProperty, lpDocument)
    '  lstrSourcePropertyValue = lobjSourceProperty.Value

    '  Try
    '    lobjPropertyPartArguments = Split(Part, ":")
    '    lstrOperation = lobjPropertyPartArguments(0).ToString.ToUpper
    '    Select Case lstrOperation
    '      Case "COMPLETE"
    '        ' Do Nothing
    '        Return lstrSourcePropertyValue
    '      Case "CONCATENATE"
    '        ' Concatenate the source property to the destination property
    '        ' Use the specified delimiter if provided
    '        Return GetConcatenatedValue(lpDocument, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL CONCATENATE"
    '        ' Concatenate the source property to the destination property only 
    '        '   if the source property currently has a certain value
    '        ' Use the specified delimiter if provided
    '        ' Expected 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:=
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:EQUALS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:<>
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EQUALS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:STARTS WITH
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:ENDS WITH
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:CONTAINS
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:NOT EMPTY
    '        '   Optional 'CONDITIONAL CONCATENATE:Delmiter:Conditional Source Value:DESTINATION EMPTY
    '        Return GetConditionalConcatenatedValue(lpDocument, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL CONCATENATE LITERAL"
    '        ' Concatenate the literal value to the destination property only 
    '        '   if the source property currently has a certain value
    '        ' Use the specified delimiter if provided
    '        'Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
    '        Return GetConditionalConcatenatedLiteral(lpDocument, lstrSourcePropertyValue, lobjPropertyPartArguments)
    '      Case "CONDITIONAL"
    '        ' Get the source property only if the destination property currently has a certain value
    '        Return GetConditionalValue(lpDocument, lobjPropertyPartArguments)
    '      Case "CONDITIONAL LITERAL"
    '        ' Apply a literal value only if the destination property currently has a certain value
    '        Return GetConditionalValue(lpDocument, lobjPropertyPartArguments, True)
    '      Case "LEFT"
    '        ' Get just the part we want from the left of the value
    '        Return Left(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '      Case "RIGHT"
    '        ' Get just the part we want from the right of the value
    '        Return Right(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '      Case "MID", "SUBSTRING"
    '        ' Get just the part we want from the middle of the value
    '        If lobjPropertyPartArguments.Length > 2 Then
    '          Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
    '        Else
    '          Return Mid(lstrSourcePropertyValue, lobjPropertyPartArguments(1))
    '        End If
    '      Case "REMOVE LEFT"
    '        ' Remove the number of characters we don't want from the left of the value
    '        Return lstrSourcePropertyValue.Remove(0, lobjPropertyPartArguments(1))
    '      Case "REMOVE RIGHT"
    '        ' Remove the number of characters we don't want from the right of the value
    '        Dim lintRightLength As Integer = CType(lobjPropertyPartArguments(1), Integer)
    '        Return lstrSourcePropertyValue.Remove((lstrSourcePropertyValue.Length - lintRightLength), lintRightLength)
    '      Case "TRIM"
    '        ' Remove the part we don't want from the left of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Both, lobjPropertyPartArguments(1))
    '      Case "TRIM LEFT"
    '        ' Remove the part we don't want from the left of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Left, lobjPropertyPartArguments(1))
    '      Case "TRIM RIGHT"
    '        ' Remove the part we don't want from the right of the value
    '        Return TrimValue(lstrSourcePropertyValue, TrimPropertyType.Right, lobjPropertyPartArguments(1))
    '      Case "REMOVE AFTER LAST INSTANCE OF CHARACTER"
    '        ' Remove everything after the last instance of the specified character
    '        ' Good for removing file extensions
    '        Return lstrSourcePropertyValue.Remove(lstrSourcePropertyValue.LastIndexOf(lobjPropertyPartArguments(1)))
    '      Case "REPLACE"
    '        ' Replace specified characters with supplied characters
    '        Return lstrSourcePropertyValue.Replace(lobjPropertyPartArguments(1), lobjPropertyPartArguments(2))
    '      Case "UPPER"
    '        ' Make all characters upper case
    '        Return lstrSourcePropertyValue.ToUpper
    '      Case "LOWER"
    '        ' Make all characters lower case
    '        Return lstrSourcePropertyValue.ToLower
    '      Case Else
    '        ' Do Nothing
    '        Return ""
    '    End Select

    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Throw New InvalidOperationException("Unable to get usable value [" & ex.Message & "]", ex)
    '  End Try

    'End Function

    ''' <summary>
    ''' Concatenates the source property to the destination property only 
    ''' if the source property currently has a certain value
    ''' </summary>
    ''' <param name="lpMetaHolder">The object from which to obtain the source value</param>
    ''' <param name="lpSourcePropertyValue">The value to check for in the source property</param>
    ''' <param name="lpPropertyPartArguments">An object array of property part arguments
    ''' Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
    ''' </param>
    ''' <returns>The concatenated literal value only if the source property currently has the specified value</returns>
    ''' <remarks>Use the specified delimiter if provided</remarks>
    Private Function GetConditionalConcatenatedLiteral(ByVal lpMetaHolder As Core.IMetaHolder,
                                                   ByVal lpSourcePropertyValue As String, ByVal lpPropertyPartArguments() As Object) As String
      Try
        ' Concatenate the source property to the destination property only 
        '   if the source property currently has a certain value
        ' Use the specified delimiter if provided
        ' Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
        Dim lstrConditionalSourceValue As String = lpPropertyPartArguments(1)
        Dim lstrLiteralValue As String = lpPropertyPartArguments(2)
        Dim lstrDelimiter As String = lpPropertyPartArguments(3)
        Dim lobjDestinationProperty As Core.ECMProperty
        Dim lstrDestinationPropertyValue As String
        Dim lstrCombinedValue As String

        lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
        lstrDestinationPropertyValue = lobjDestinationProperty.Value

        If lpSourcePropertyValue = lstrConditionalSourceValue Then
          lstrCombinedValue = lstrDestinationPropertyValue & lstrDelimiter & lstrLiteralValue
          Return lstrCombinedValue
        Else
          ' The condition was not met, just return the original destination value
          Return lstrDestinationPropertyValue
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function
    'Private Function GetConditionalConcatenatedLiteral(ByVal lpDocument As Core.Document, _
    '                                                   ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String


    'End Function

    'Private Function GetConditionalConcatenatedLiteral(ByVal lpVersion As Core.Version, _
    '                                                   ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String

    '  ' Concatenate the source property to the destination property only 
    '  '   if the source property currently has a certain value
    '  ' Use the specified delimiter if provided
    '  ' Expected 'CONDITIONAL CONCATENATE LITERAL:Conditional Source Value:Literal Value:Delimiter
    '  Dim lstrConditionalSourceValue As String = lobjPropertyPartArguments(1)
    '  Dim lstrLiteralValue As String = lobjPropertyPartArguments(2)
    '  Dim lstrDelimiter As String = lobjPropertyPartArguments(3)
    '  Dim lobjDestinationProperty As Core.ECMProperty
    '  Dim lstrDestinationPropertyValue As String
    '  Dim lstrCombinedValue As String

    '  lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '  lstrDestinationPropertyValue = lobjDestinationProperty.Value

    '  If lpSourcePropertyValue = lstrConditionalSourceValue Then
    '    lstrCombinedValue = lstrDestinationPropertyValue & lstrDelimiter & lstrLiteralValue
    '    Return lstrCombinedValue
    '  Else
    '    ' The condition was not met, just return the original destination value
    '    Return lstrDestinationPropertyValue
    '  End If

    'End Function

    ''' <summary>
    ''' Concatenates the source property to the destination property only 
    ''' if the source property currently has a certain value
    ''' </summary>
    ''' <param name="lpMetaHolder"></param>
    ''' <param name="lpSourcePropertyValue"></param>
    ''' <param name="lpPropertyPartArguments">Expected CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:=
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EQUALS
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:STARTS WITH
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:ENDS WITH
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:CONTAINS
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EMPTY
    '''   Optional CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:DESTINATION EMPTY
    ''' </param>
    ''' <returns></returns>
    ''' <remarks>Uses the specified delimiter if provided</remarks>
    Private Function GetConditionalConcatenatedValue(ByVal lpMetaHolder As Core.IMetaHolder,
                                                 ByVal lpSourcePropertyValue As String, ByVal lpPropertyPartArguments() As Object) As String

      ' Concatenate the source property to the destination property only 
      '   if the source property currently has a certain value
      ' Use the specified delimiter if provided
      ' Expected 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:=
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EQUALS
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:STARTS WITH
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:ENDS WITH
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:CONTAINS
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EMPTY
      '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:DESTINATION EMPTY

      Dim lobjDestinationProperty As Core.ECMProperty
      Dim lstrDestinationPropertyValue As String
      Dim lstrConditionalSourceValue As String = lpPropertyPartArguments(2)

      ' Check to see if there is a part defining the operator, otherwise default to equals
      If lpPropertyPartArguments.Length > 3 Then
        Dim lstrConditionalOperator As String = lpPropertyPartArguments(3)
        Select Case lstrConditionalOperator.Trim.ToUpper
          Case "=", "EQUALS"
            If lpSourcePropertyValue.ToUpper = lstrConditionalSourceValue.ToUpper Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case "<>", "NOT EQUALS"
            If lpSourcePropertyValue.ToUpper <> lstrConditionalSourceValue.ToUpper Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case "STARTS WITH"
            If lpSourcePropertyValue.ToUpper.StartsWith(lstrConditionalSourceValue.ToUpper) Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case "ENDS WITH"
            If lpSourcePropertyValue.ToUpper.EndsWith(lstrConditionalSourceValue.ToUpper) Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case "CONTAINS"
            If lpSourcePropertyValue.ToUpper.Contains(lstrConditionalSourceValue.ToUpper) Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case "NOT EMPTY"
            If lpSourcePropertyValue.Length > 0 Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case ""
            lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
            lstrDestinationPropertyValue = lobjDestinationProperty.Value
            If lstrDestinationPropertyValue.Length = 0 Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

          Case Else
            If lpSourcePropertyValue = lstrConditionalSourceValue Then
              Return GetConcatenatedValue(lpMetaHolder, lpSourcePropertyValue, lpPropertyPartArguments)
            Else
              lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
              lstrDestinationPropertyValue = lobjDestinationProperty.Value
              Return lstrDestinationPropertyValue
            End If

        End Select

      Else
        ' No operator was defined assume an operator of equals
        lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
        lstrDestinationPropertyValue = lobjDestinationProperty.Value
        Return lstrDestinationPropertyValue

      End If

      ' The condition was not met, just return the original destination value
      'Dim lobjDestinationProperty As ECMProperty
      'Dim lstrDestinationPropertyValue As String
      'lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
      'lstrDestinationPropertyValue = lobjDestinationProperty.Value
      'Return lstrDestinationPropertyValue
      'End If

    End Function

    'Private Function GetConditionalConcatenatedValue(ByVal lpDocument As Core.Document, _
    '                                                 ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String

    '  ' Concatenate the source property to the destination property only 
    '  '   if the source property currently has a certain value
    '  ' Use the specified delimiter if provided
    '  ' Expected 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:=
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EQUALS
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:STARTS WITH
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:ENDS WITH
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:CONTAINS
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EMPTY
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:DESTINATION EMPTY

    '  Dim lobjDestinationProperty As Core.ECMProperty
    '  Dim lstrDestinationPropertyValue As String
    '  Dim lstrConditionalSourceValue As String = lobjPropertyPartArguments(2)

    '  ' Check to see if there is a part defining the operator, otherwise default to equals
    '  If lobjPropertyPartArguments.Length > 3 Then
    '    Dim lstrConditionalOperator As String = lobjPropertyPartArguments(3)
    '    Select Case lstrConditionalOperator.Trim.ToUpper
    '      Case "=", "EQUALS"
    '        If lpSourcePropertyValue.ToUpper = lstrConditionalSourceValue.ToUpper Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "<>", "NOT EQUALS"
    '        If lpSourcePropertyValue.ToUpper <> lstrConditionalSourceValue.ToUpper Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "STARTS WITH"
    '        If lpSourcePropertyValue.ToUpper.StartsWith(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "ENDS WITH"
    '        If lpSourcePropertyValue.ToUpper.EndsWith(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "CONTAINS"
    '        If lpSourcePropertyValue.ToUpper.Contains(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "NOT EMPTY"
    '        If lpSourcePropertyValue.Length > 0 Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case ""
    '        lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '        lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '        If lstrDestinationPropertyValue.Length = 0 Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case Else
    '        If lpSourcePropertyValue = lstrConditionalSourceValue Then
    '          Return GetConcatenatedValue(lpDocument, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '    End Select

    '  Else
    '    ' No operator was defined assume an operator of equals
    '    lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '    lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '    Return lstrDestinationPropertyValue

    '  End If

    '  ' The condition was not met, just return the original destination value
    '  'Dim lobjDestinationProperty As ECMProperty
    '  'Dim lstrDestinationPropertyValue As String
    '  'lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '  'lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '  'Return lstrDestinationPropertyValue
    '  'End If

    'End Function

    'Private Function GetConditionalConcatenatedValue(ByVal lpVersion As Core.Version, _
    '                                                 ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String

    '  ' Concatenate the source property to the destination property only 
    '  '   if the source property currently has a certain value
    '  ' Use the specified delimiter if provided
    '  ' Expected 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:=
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EQUALS
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:STARTS WITH
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:ENDS WITH
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:CONTAINS
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:NOT EMPTY
    '  '   Optional 'CONDITIONAL CONCATENATE:Delimiter:Conditional Source Value:DESTINATION EMPTY

    '  Dim lobjDestinationProperty As Core.ECMProperty
    '  Dim lstrDestinationPropertyValue As String
    '  Dim lstrConditionalSourceValue As String = lobjPropertyPartArguments(2)

    '  ' Check to see if there is a part defining the operator, otherwise default to equals
    '  If lobjPropertyPartArguments.Length > 3 Then
    '    Dim lstrConditionalOperator As String = lobjPropertyPartArguments(3)
    '    Select Case lstrConditionalOperator.Trim.ToUpper
    '      Case "=", "EQUALS"
    '        If lpSourcePropertyValue.ToUpper = lstrConditionalSourceValue.ToUpper Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "<>", "NOT EQUALS"
    '        If lpSourcePropertyValue.ToUpper <> lstrConditionalSourceValue.ToUpper Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "STARTS WITH"
    '        If lpSourcePropertyValue.ToUpper.StartsWith(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "ENDS WITH"
    '        If lpSourcePropertyValue.ToUpper.EndsWith(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "CONTAINS"
    '        If lpSourcePropertyValue.ToUpper.Contains(lstrConditionalSourceValue.ToUpper) Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case "NOT EMPTY"
    '        If lpSourcePropertyValue.Length > 0 Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case ""
    '        lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '        lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '        If lstrDestinationPropertyValue.Length = 0 Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '      Case Else
    '        If lpSourcePropertyValue = lstrConditionalSourceValue Then
    '          Return GetConcatenatedValue(lpVersion, lpSourcePropertyValue, lobjPropertyPartArguments)
    '        Else
    '          lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '          lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '          Return lstrDestinationPropertyValue
    '        End If

    '    End Select

    '  Else
    '    ' No operator was defined assume an operator of equals
    '    lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '    lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '    Return lstrDestinationPropertyValue

    '  End If

    '  ' The condition was not met, just return the original destination value
    '  'Dim lobjDestinationProperty As ECMProperty
    '  'Dim lstrDestinationPropertyValue As String
    '  'lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '  'lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '  'Return lstrDestinationPropertyValue
    '  'End If

    'End Function

    ''' <summary>
    ''' Concatenates the source property to the destination property
    ''' </summary>
    ''' <param name="lpMetaHolder">The object from which to obtain the source value</param>
    ''' <param name="lpSourcePropertyValue">The value to concatenate to</param>
    ''' <param name="lpPropertyPartArguments">An object array of property part arguments</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetConcatenatedValue(ByVal lpMetaHolder As Core.IMetaHolder, ByVal lpSourcePropertyValue As String, ByVal lpPropertyPartArguments() As Object) As String
      Try
        ' Concatenate the source property to the destination property
        ' Use the specified delimiter if provided
        Dim lobjDestinationProperty As Core.ECMProperty
        Dim lstrDestinationPropertyValue As String
        Dim lstrDelimiter As String = lpPropertyPartArguments(1)
        ' Dim lstrCombinedValue As String
        lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
        lstrDestinationPropertyValue = lobjDestinationProperty.Value

        ' <Modified by: Ernie at 1/10/2013-2:57:40 PM on machine: ERNIE-THINK>
        ' lstrCombinedValue = lstrDestinationPropertyValue & lstrDelimiter & lpSourcePropertyValue
        ' Return lstrCombinedValue
        If Not String.IsNullOrEmpty(lpSourcePropertyValue) Then
          Return String.Format("{0}{1}{2}", lstrDestinationPropertyValue, lstrDelimiter, lpSourcePropertyValue)
        Else
          Return lstrDestinationPropertyValue
        End If
        ' </Modified by: Ernie at 1/10/2013-2:57:40 PM on machine: ERNIE-THINK>

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <Added by: Ernie at: 9/10/2013-10:51:19 AM on machine: ERNIE-THINK>
    Private Function GetLiteralConcatenatedValue(ByVal lpMetaHolder As Core.IMetaHolder, ByVal lpPropertyPartArguments() As Object) As String
      Try
        Dim lobjDestinationProperty As Core.ECMProperty
        Dim lstrDestinationPropertyValue As String

        Select Case lpPropertyPartArguments.Length
          Case 0
            Throw New ArgumentNullException("lpPropertyPartArguments")

          Case 3
            Dim lstrDelimiter As String = lpPropertyPartArguments(1)
            Dim lstrValueToConcatenate As String = lpPropertyPartArguments(2)
            lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)
            lstrDestinationPropertyValue = lobjDestinationProperty.Value

            lstrDestinationPropertyValue = String.Format("{0}{1}{2}", lstrDestinationPropertyValue, lstrDelimiter, lstrValueToConcatenate)

            Return lstrDestinationPropertyValue

          Case Else
            Throw New ArgumentOutOfRangeException("lpPropertyPartArguments",
              String.Format("Property Part Arguments contains {0} elements, 3 elements were expected.", lpPropertyPartArguments.Length))

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function
    ' </Added by: Ernie at: 9/10/2013-10:51:19 AM on machine: ERNIE-THINK>

    Private Function GetPrefixedValue(ByVal lpMetaHolder As Core.IMetaHolder, ByVal lpSourcePropertyValue As String, ByVal lpPropertyPartArguments() As Object) As String
      Try
        ' Concatenate the source property to the destination property
        ' Use the specified delimiter if provided
        Dim lobjDestinationProperty As Core.ECMProperty
        Dim lstrDestinationPropertyValue As String
        Dim lstrDelimiter As String = lpPropertyPartArguments(1)
        Dim lstrCombinedValue As String
        lobjDestinationProperty = GetProperty(DestinationProperty, lpMetaHolder)

        lstrDestinationPropertyValue = lobjDestinationProperty.Value

        lstrCombinedValue = lpSourcePropertyValue & lstrDelimiter & lstrDestinationPropertyValue
        Return lstrCombinedValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Function GetConcatenatedValue(ByVal lpDocument As Core.Document, ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String

    '  ' Concatenate the source property to the destination property
    '  ' Use the specified delimiter if provided
    '  Dim lobjDestinationProperty As Core.ECMProperty
    '  Dim lstrDestinationPropertyValue As String
    '  Dim lstrDelimiter As String = lobjPropertyPartArguments(1)
    '  Dim lstrCombinedValue As String
    '  lobjDestinationProperty = GetProperty(DestinationProperty, lpDocument)
    '  lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '  lstrCombinedValue = lstrDestinationPropertyValue & lstrDelimiter & lpSourcePropertyValue
    '  Return lstrCombinedValue

    'End Function

    'Private Function GetConcatenatedValue(ByVal lpVersion As Core.Version, ByVal lpSourcePropertyValue As String, ByVal lobjPropertyPartArguments() As Object) As String

    '  ' Concatenate the source property to the destination property
    '  ' Use the specified delimiter if provided
    '  Dim lobjDestinationProperty As Core.ECMProperty
    '  Dim lstrDestinationPropertyValue As String
    '  Dim lstrDelimiter As String = lobjPropertyPartArguments(1)
    '  Dim lstrCombinedValue As String
    '  lobjDestinationProperty = GetProperty(DestinationProperty, lpVersion)
    '  lstrDestinationPropertyValue = lobjDestinationProperty.Value
    '  lstrCombinedValue = lstrDestinationPropertyValue & lstrDelimiter & lpSourcePropertyValue
    '  Return lstrCombinedValue

    'End Function

    Private Shared Function TrimValue(ByVal lpValue As String,
                               ByVal lpTrimType As TrimPropertyType,
                               Optional ByVal lpTrimCharacter As String = " ") As String

      Select Case lpTrimType
        Case TrimPropertyType.Left
          Return lpValue.TrimStart(lpTrimCharacter)
        Case TrimPropertyType.Right
          Return lpValue.TrimEnd(lpTrimCharacter)
        Case TrimPropertyType.Both
          Return lpValue.Trim(lpTrimCharacter)
        Case Else
          Return lpValue
      End Select

    End Function

    'Private Function RemoveAfterLastInstanceOfCharacter(ByVal lpValue As String, _
    '  ByVal lpRemovalCharacterKey As String) As String

    '  Dim lintKeyCharacterPosition As integer
    '  lintKeyCharacterPosition = lpValue.LastIndexOf(lpRemovalCharacterKey)
    '  Return lpValue.Remove(lintKeyCharacterPosition)

    'End Function

    ' ''Public Function TrimProperty(ByVal lpPropertyScope As PropertyScope, _
    ' ''ByVal lpPropertyName As String, _
    ' ''ByVal lpTrimType As TrimPropertyValue.TrimPropertyType, _
    ' ''Optional ByVal lpTrimCharacter As String = " ") As Boolean

    ' ''  Dim lstrPropertyValue As String
    ' ''  Select Case lpPropertyScope
    ' ''    Case PropertyScope.VersionProperty
    ' ''      Try
    ' ''        For Each lobjVersion As Version In Versions
    ' ''          lstrPropertyValue = lobjVersion.Properties(lpPropertyName).Value
    ' ''          lstrPropertyValue = TrimValue(lstrPropertyValue, lpTrimType, lpTrimCharacter)
    ' ''          lobjVersion.Properties(lpPropertyName).Value = lstrPropertyValue
    ' ''        Next
    ' ''        Return True
    ' ''      Catch ex As Exception
    ' ''        Return False
    ' ''      End Try

    ' ''    Case PropertyScope.DocumentProperty
    ' ''      lstrPropertyValue = Properties(lpPropertyName).Value
    ' ''      lstrPropertyValue = TrimValue(lstrPropertyValue, lpTrimType, lpTrimCharacter)
    ' ''      Properties(lpPropertyName).Value = lstrPropertyValue

    ' ''  End Select
    ' ''End Function

#End Region

#Region "DataLookup Implementation"

    Public Overrides Function SourceExists(ByVal lpMetaHolder As Core.IMetaHolder) As Boolean
      Try
        If Me.SourceProperty IsNot Nothing AndAlso lpMetaHolder.Metadata.PropertyExists(SourceProperty.PropertyName, False) Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets the value from the source property.
    ''' </summary>
    ''' <param name="lpMetaHolder">The object to get the value from.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function GetValue(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Try

        Return GetUsableValue(lpMetaHolder)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Overrides Function GetValues(ByVal lpMetaHolder As Core.IMetaHolder) As Object
      Return Nothing
    End Function

    '''' <summary>Gets the value from the source property.</summary>
    '''' <param name="lpDocument">The document to get the value from.</param>
    'Public Overrides Function GetValue(ByVal lpDocument As Core.Document) As Object
    '  Return GetUsableValue(lpDocument)
    'End Function

    'Public Overrides Function GetValue(ByVal lpVersion As Core.Version) As Object
    '  Return GetUsableValue(lpVersion)
    'End Function

    Public Overrides Function GetParameters() As IParameters Implements IPropertyLookup.GetParameters
      Try
        Dim lobjParameters As IParameters = New Parameters

        If SourceProperty Is Nothing Then
          Throw New InvalidOperationException("The source property is not initialized.")
        End If

        If DestinationProperty Is Nothing Then
          Throw New InvalidOperationException("The destination property is not initialized.")
        End If

        Dim lstrFirstPart As String
        Dim lstrLastPart As String = String.Empty

        If Part.Contains(":") Then
          lstrFirstPart = Part.Substring(0, Part.IndexOf(":")).Replace(" ", String.Empty)
          lstrLastPart = Part.Substring(Part.IndexOf(":") + 1)
        Else
          lstrFirstPart = Part.Replace(" ", String.Empty)
        End If

        Dim lenuParseStyle As PartEnum = GetPartStyle(Part, lstrLastPart)

        If lobjParameters.Contains(PARAM_PARSE_STYLE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PARSE_STYLE, lenuParseStyle, GetType(PartEnum),
            "Specifies the style of parsing to use for changing the property value."))
        End If

        If lobjParameters.Contains(PARAM_PARSE_ARGUMENTS) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PARSE_ARGUMENTS, lstrLastPart,
            "(Optional) Additional arguments for the parse style."))
        End If

        For Each lobjLookupPropertyParameter As IParameter In SourceProperty.GetParameters()
          lobjLookupPropertyParameter.Name = String.Format("Source{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next
        For Each lobjLookupPropertyParameter As IParameter In DestinationProperty.GetParameters()
          lobjLookupPropertyParameter.Name = String.Format("Destination{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Shared Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_GENERATE_CDF) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_GENERATE_CDF, True,
            "Specifies whether or not the exported document should be saved to a file."))
        End If

        If lobjParameters.Contains(PARAM_PARSE_STYLE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PARSE_STYLE, PartEnum.Complete, GetType(PartEnum),
            "Specifies the style of parsing to use for changing the property value."))
        End If

        If lobjParameters.Contains(PARAM_PARSE_ARGUMENTS) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PARSE_ARGUMENTS, String.Empty,
            "(Optional) Additional arguments for the parse style."))
        End If

        For Each lobjLookupPropertyParameter As IParameter In LookupProperty.GetDefaultParameters()
          lobjLookupPropertyParameter.Name = String.Format("Source{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next
        For Each lobjLookupPropertyParameter As IParameter In LookupProperty.GetDefaultParameters()
          lobjLookupPropertyParameter.Name = String.Format("Destination{0}", lobjLookupPropertyParameter.Name)
          lobjParameters.Add(lobjLookupPropertyParameter)
        Next

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace