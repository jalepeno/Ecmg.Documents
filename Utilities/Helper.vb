Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Utilities

  Public Class Helper

    ''' <summary>
    '''     Takes an incoming SQL Select statement and converts it into a SELECT COUNT(*) statement.
    ''' </summary>
    ''' <param name="lpSQL" type="String">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     A SELECT COUNT(*) statement.
    ''' </returns>
    Public Shared Function ConvertSelectSQLToSelectCount(lpSQL As String) As String
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpSQL)
#Else
          If lpSQL Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpSQL))
          End If
#End If

        Dim lstrSQL As String = lpSQL.ToLower

        If Not lstrSQL.Contains("select") Then
          Throw New InvalidOperationException("Incoming SQL statement is not a SELECT statement.")
        End If

        Dim lintFirstSpaceAfterSelect As Integer = lstrSQL.IndexOf("select ") + 7
        Dim lintFromPosition As Integer = lstrSQL.IndexOf(" from")
        Return String.Format("{0} Count(*) {1}", lpSQL.Substring(0, lintFirstSpaceAfterSelect), lpSQL.Substring(lintFromPosition))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function
    Public Shared Function CopyByteArrayToStream(lpBytes As Byte()) As Stream
      Try

        Return New MemoryStream(lpBytes)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Shared Function FormatPercentage(lpPercentage As Single) As String
      Try
        Return lpPercentage.ToString("P")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function FormatTimeSpan(lpTimeSpan As TimeSpan) As String
      Try
        Dim lobjStringBuilder As New StringBuilder

        If lpTimeSpan.TotalSeconds < 1 Then
          lobjStringBuilder.AppendFormat("{0} Milliseconds", lpTimeSpan.Milliseconds)
        ElseIf lpTimeSpan.TotalMinutes < 1 Then
          lobjStringBuilder.AppendFormat("{0}.{1} Seconds", lpTimeSpan.Seconds, lpTimeSpan.Milliseconds)
        ElseIf lpTimeSpan.TotalHours < 1 Then
          lobjStringBuilder.AppendFormat("{0} Minutes {1} Seconds", lpTimeSpan.Minutes, lpTimeSpan.Seconds)
        ElseIf lpTimeSpan.TotalDays < 1 Then
          lobjStringBuilder.AppendFormat("{0} Hours {1} Minutes {2} Seconds", lpTimeSpan.Hours, lpTimeSpan.Minutes, lpTimeSpan.Seconds)
        Else
          lobjStringBuilder.AppendFormat("{0} Days {1} Hours {2} Minutes {3} Seconds", lpTimeSpan.Days, lpTimeSpan.Hours, lpTimeSpan.Minutes, lpTimeSpan.Seconds)
        End If

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Ping(lpHostNameOrAddress As String, lpTimeout As Integer) As Boolean
      Try
        Using lobjPingSender As New Ping
          ' Use the default Ttl value which Is 128,
          ' but change the fragmentation behavior.
          Dim lobjPingOptions As New PingOptions With {.DontFragment = True}

          ' Create a buffer of 32 bytes of data to be transmitted.
          Dim lstrData As String = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
          Dim lobjBuffer As Byte() = Encoding.ASCII.GetBytes(lstrData)

          Dim lobjReply As PingReply = lobjPingSender.Send(lpHostNameOrAddress, lpTimeout, lobjBuffer, lobjPingOptions)
          If lobjReply.Status = IPStatus.Success Then
            Return True
          Else
            Return False
          End If
        End Using
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Converts the string representation of the name or numeric value 
    ''' of one or more enumerated constants to an equivalent enumerated object.
    ''' </summary>
    ''' <param name="lpString">A string containing the name or value to convert.</param>
    ''' <param name="lpType">The System.Type of the enumeration.</param>
    ''' <returns>An object of type enumType whose value is represented by value.</returns>
    ''' <remarks></remarks>
    Public Shared Function StringToEnum(ByVal lpString As String,
                                ByVal lpType As Type) As Object

      Try

        Dim lenuObject As Object = [Enum].Parse(lpType, lpString)

        Return lenuObject

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        Return 0
      End Try

    End Function

    Public Shared Sub AssignObjectProperty(lpPropertyName As String, ByVal lpSourceObject As Object, ByVal lpDestinationObject As Object)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpSourceObject)
        ArgumentNullException.ThrowIfNull(lpDestinationObject)
#Else
        If lpSourceObject Is Nothing Then
          Throw New ArgumentNullException(nameof(lpSourceObject))
        End If
        If lpDestinationObject Is Nothing Then
          Throw New ArgumentNullException(nameof(lpDestinationObject))
        End If
#End If

        'Dim lobjSourceProperties() As Reflection.PropertyInfo = lpSourceObject.GetType.GetProperties
        'Dim lobjDestinationProperties() As Reflection.PropertyInfo = lpDestinationObject.GetType.GetProperties
        Dim lobjSourcePropertyInfo As Reflection.PropertyInfo
        Dim lobjDestinationPropertyInfo As Reflection.PropertyInfo
        Dim lobjSourcePropertyValue As Object = Nothing
        Dim lobjSourceParentValue As Object = Nothing
        Dim lobjDestinationPropertyValue As Object = Nothing
        Dim lobjDestinationParentValue As Object = Nothing
        Dim lobjObjectValue As Object = Nothing

        lobjSourcePropertyInfo = GetPropertyInfo(lpPropertyName, lpSourceObject, lobjSourcePropertyValue, lobjSourceParentValue)
        lobjDestinationPropertyInfo = GetPropertyInfo(lpPropertyName, lpDestinationObject, lobjDestinationPropertyValue, lobjDestinationParentValue)

        If lobjSourcePropertyInfo.CanRead Then
          'lobjObjectValue = lobjSourcePropertyInfo.GetValue(lpSourceObject, Nothing)
          lobjObjectValue = lobjSourcePropertyValue
        Else
          Throw New InvalidOperationException(String.Format("The source property '{0}' can't be read.", lpPropertyName))
        End If
        'lobjSourcePropertyInfo.
        If lobjDestinationPropertyInfo.CanWrite = True Then
          If lpPropertyName.Contains("."c) Then
            lobjDestinationPropertyInfo.SetValue(lobjDestinationParentValue, lobjObjectValue, Nothing)
          Else
            lobjDestinationPropertyInfo.SetValue(lpDestinationObject, lobjObjectValue, Nothing)
          End If
          ''          Dim lobjValueType As Type = lobjDestinationPropertyInfo.PropertyType
          'CType(lobjObjectValue, lobjValueType.GetType.Name)
          'DirectCast(lobjObjectValue, lobjValueType.Name)
          'lobjObjectValue = Convert.ChangeType(lobjObjectValue, lobjValueType)
          ''        lobjDestinationPropertyInfo.SetValue(lobjDestinationParentValue, lobjObjectValue, Nothing)
        Else
          Throw New InvalidOperationException(String.Format("The destination property '{0}' can't be written.", lpPropertyName))
        End If

      Catch ex As Exception

        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Function GetPropertyInfo(lpPropertyName As String, ByVal lpObject As Object, ByRef lpPropertyValue As Object, ByRef lpParentValue As Object) As Reflection.PropertyInfo
      Try

        Dim lobjPropertyInfo As Reflection.PropertyInfo = Nothing

        If lpPropertyName.Contains("."c) Then
          Dim lstrPropertyChain As String() = lpPropertyName.Split(".")
          Dim lobjCurrentObjectLevel As Object = lpObject
          Dim lobjPreviousObjectLevel As Object = Nothing
          For Each lstrPropertyLevel As String In lstrPropertyChain
            lobjPropertyInfo = lobjCurrentObjectLevel.GetType.GetProperty(lstrPropertyLevel)
            If lobjPropertyInfo Is Nothing Then
              Throw New InvalidOperationException(String.Format("No property exists with the name '{0}'.", lstrPropertyLevel))
            End If
            lobjPreviousObjectLevel = lobjCurrentObjectLevel
            lobjCurrentObjectLevel = lobjPropertyInfo.GetValue(lobjCurrentObjectLevel, Nothing)
            lpPropertyValue = lobjCurrentObjectLevel
            lpParentValue = lobjPreviousObjectLevel
          Next
        Else
          lobjPropertyInfo = lpObject.GetType.GetProperty(lpPropertyName)
          If lobjPropertyInfo Is Nothing Then
            Throw New InvalidOperationException(String.Format("No property exists with the name '{0}'.", lpPropertyName))
          End If
          lpPropertyValue = lobjPropertyInfo.GetValue(lpObject, Nothing)
        End If

        Return lobjPropertyInfo

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Shared Sub WriteConsoleErrorMessage(lpMessage As String)
      Try
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine(lpMessage)
        Console.ResetColor()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    Public Shared Function WriteFileToByteArray(ByVal lpFilePath As String) As Byte()
      Try
        Return CopyStreamToByteArray(WriteFileToMemoryStream(lpFilePath))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Sets the top level properties of the destination object based on the top level properties of the source object
    ''' </summary>
    ''' <param name="lpSourceObject">The object containing the property values to be used</param>
    ''' <param name="lpDestinationObject">The object whose property values are to be set</param>
    ''' <remarks>Properties are matched up by property name</remarks>
    Public Shared Sub AssignObjectProperties(ByVal lpSourceObject As Object, ByVal lpDestinationObject As Object)

      Try

        If lpSourceObject IsNot Nothing Then

#If NET8_0_OR_GREATER Then
          ArgumentNullException.ThrowIfNull(lpDestinationObject)
#Else
          If lpDestinationObject Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpDestinationObject))
          End If
#End If

          Dim lobjSourceProperties() As Reflection.PropertyInfo = lpSourceObject.GetType.GetProperties
          Dim lobjDestinationProperties() As Reflection.PropertyInfo = lpDestinationObject.GetType.GetProperties
          Dim lobjSourcePropertyInfo As Reflection.PropertyInfo
          Dim lobjDestinationPropertyInfo As Reflection.PropertyInfo
          Dim lobjObjectValue As Object

          For Each lobjSourcePropertyInfo In lobjSourceProperties
            Try
              If lobjSourcePropertyInfo.CanRead Then
                lobjObjectValue = lobjSourcePropertyInfo.GetValue(lpSourceObject, Nothing)
              Else
                ' This property is write-only or for some other reason may not be read, skip over it.
                Continue For
              End If
            Catch CountException As Reflection.TargetParameterCountException
              '  This is a parameter count mismatch, skip this property and move on...
              Continue For
            End Try

            For Each lobjDestinationPropertyInfo In lobjDestinationProperties
              If lobjDestinationPropertyInfo.Name = lobjSourcePropertyInfo.Name AndAlso lobjDestinationPropertyInfo.CanWrite = True Then
                lobjDestinationPropertyInfo.SetValue(lpDestinationObject, lobjObjectValue, Nothing)
                Exit For
              End If
            Next
          Next
        Else
          ' We can't assign the properties from the source if it is nothing
          Throw New ArgumentNullException(NameOf(lpSourceObject))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

    Public Shared Function CopyStreamToTempFileStream(lpInputStream As Stream, ByRef lpFileName As String) As FileStream
      Try
        lpFileName = Path.GetTempFileName()
        Helper.WriteStreamToFile(lpInputStream, lpFileName, FileMode.Create)
        Return New FileStream(lpFileName, FileMode.Open)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub WriteStreamToFile(lpInputStream As Stream, lpOutputPath As String, lpMode As FileMode)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpInputStream)
#Else
          If lpInputStream Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpInputStream))
          End If
#End If

        If lpInputStream.CanRead = False Then
          Throw New InvalidOperationException("The stream is not readable")
        End If
        Dim lobjOutputStream As New FileStream(lpOutputPath, lpMode)
        If lpInputStream.CanSeek Then
          lpInputStream.Seek(0, 0)
        End If

        'lpInputStream.CopyTo(lobjOutputStream)

        CopyStream(lpInputStream, lobjOutputStream)

        lobjOutputStream.Close()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function ObjectContainsProperty(ByVal lpObject As Object, ByVal lpPropertyName As String) As Boolean
      Try
        Dim lobjObjectType As Type = lpObject.GetType

        Return ObjectTypeContainsProperty(lobjObjectType, lpPropertyName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ObjectTypeContainsProperty(ByVal lpType As Type, ByVal lpPropertyName As String) As Boolean
      Try

        For Each lobjPropertyInfo As PropertyInfo In lpType.GetProperties
          If String.Compare(lpPropertyName, lobjPropertyInfo.Name) = 0 Then
            Return True
          End If
        Next

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determines whether or not a the current method 
    ''' was called as part of a deserialization process
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsDeserializationBasedCall() As Boolean
      Try
        Return CallStackContainsMethodName("Deserialize", "LoadFromXmlDocument")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determines whether or not a the current method 
    ''' was called as part of a serialization process
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsSerializationBasedCall() As Boolean
      Try
        Return CallStackContainsMethodName("Serialize", "Save", "ToXmlString")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determines whether or not the call stack contains any of the specified method names
    ''' </summary>
    ''' <param name="lpMethodNames">One or more method names to check</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CallStackContainsMethodName(ByVal ParamArray lpMethodNames() As String) As Boolean
      Try
        If lpMethodNames.Length <= 0 Then Return False ' No arguments passed.
        Dim lobjStackTrace As New System.Diagnostics.StackTrace()
        Dim lstrCallingMethod As String
        Dim lstrSmallMethodName As String

        For lintStackFrameCounter As Integer = 1 To lobjStackTrace.FrameCount - 1
          With lobjStackTrace.GetFrame(lintStackFrameCounter)
            lstrCallingMethod = .GetMethod.ToString.ToLower
            For lintMethodNameCounter As Integer = 0 To UBound(lpMethodNames, 1)
              lstrSmallMethodName = lpMethodNames(lintMethodNameCounter).ToLower
              If lstrCallingMethod.Contains(lstrSmallMethodName) Then
                Return True
              End If
            Next lintMethodNameCounter
          End With
        Next

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateDisplayName(lpName As String) As String
      Try
        Return AddSpacesToSentence(lpName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function AddSpacesToSentence(lpOriginalText As String) As String
      Try
        If IsNullOrWhiteSpace(lpOriginalText) Then
          Return String.Empty
        End If
        Dim lobjStringBuilder As New StringBuilder(lpOriginalText.Length * 2)
        lobjStringBuilder.Append(lpOriginalText(0))
        For lintCounter As Integer = 1 To (lpOriginalText.Length - 1)
          If Char.IsUpper(lpOriginalText(lintCounter)) AndAlso lpOriginalText(lintCounter - 1) <> " " AndAlso Char.IsLower(lpOriginalText(lintCounter - 1)) Then
            lobjStringBuilder.Append(" "c)
          End If
          lobjStringBuilder.Append(lpOriginalText(lintCounter))
        Next

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks the specified stream to see if it represents a modca file.
    ''' </summary>
    ''' <param name="lpStream">The stream to check.</param>
    ''' <returns>True if the first line contains
    ''' a MO:DCA header, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsModcaStream(ByVal lpStream As Stream) As Boolean
      Try

        Dim lblnReturnValue As Boolean

        If (lpStream.CanSeek = False) Then
          Throw New Exception("IsModcaStream requires a stream that is seekable.")
        End If

        ' TODO: Modify this code to identify the MODCA magic numbers instead of the TIFF magic numbers.
        ' https://en.wikipedia.org/wiki/MODCA
        ' Magic number	X'D3', X'D3A8', X'D3A9'

        ' From email from Neil Peckman dated 1/6/2016
        ' The 1st eight bytes look Like this in Hex: 

        ' 12345678 
        ' ..Lyy... 
        ' 02DAA000 
        ' 01388000 

        ' In ASCII it looks like this: 

        ' .!Ë¿¿... 
        ' 02DAA000 
        ' 01388000 

        ' The "is_MODCA" test would be true when  char 3,4 = X" D300", X"D3A8", or X"D3A9",   
        ' (that would be the 'magic number')  - https://en.wikipedia.org/wiki/MODCA according to wiki.  
        ' Pretty much all you'll ever see is x'D3A8A8'.  The 1st 2 bytes are a length that we shouldn't bother with. 

        Const HEADER_SIZE As Integer = 4
        Const EMPTY_BYTE As Byte = 0
        Const MODCA_MAGIC_NUMBER_1 As Byte = 211  ' HEX D3
        Const MODCA_MAGIC_NUMBER_2 As Byte = 168  ' HEX A8
        Const MODCA_MAGIC_NUMBER_3 As Byte = 169  ' HEX A9

        lpStream.Seek(0, SeekOrigin.Begin)

        If lpStream.Length < HEADER_SIZE Then
          Return False
        End If

        Dim lobjHeader As Byte() = New Byte(HEADER_SIZE) {}

        lpStream.Read(lobjHeader, 0, lobjHeader.Length)

        If ((lobjHeader(2) = MODCA_MAGIC_NUMBER_1) AndAlso (lobjHeader(3) = EMPTY_BYTE)) Then
          lblnReturnValue = True
        ElseIf ((lobjHeader(2) = MODCA_MAGIC_NUMBER_1) AndAlso (lobjHeader(3) = MODCA_MAGIC_NUMBER_2)) Then
          lblnReturnValue = True
        ElseIf ((lobjHeader(2) = MODCA_MAGIC_NUMBER_1) AndAlso (lobjHeader(3) = MODCA_MAGIC_NUMBER_3)) Then
          lblnReturnValue = True
        End If

        lpStream.Seek(0, SeekOrigin.Begin)

        Return lblnReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function IsNullOrWhiteSpace(lpValue As String) As Boolean
      Try
        If lpValue Is Nothing Then
          Return True
        End If

        For i As Integer = 0 To lpValue.Length - 1
          If Not Char.IsWhiteSpace(lpValue(i)) Then
            Return False
          End If
        Next

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Pass through method for WriteExceptionInfo
    ''' </summary>
    ''' <param name="Ex">The exception to dunp</param>
    ''' <remarks></remarks>
    Public Shared Sub DumpException(ByVal Ex As Exception)

      Try
        WriteExceptionInfo(Ex)
        If Ex.InnerException IsNot Nothing Then
          WriteExceptionInfo(Ex.InnerException)
        End If
      Catch newEx As Exception
        ApplicationLogging.LogException(Ex, MethodBase.GetCurrentMethod)
      End Try

    End Sub

    Public Shared Function EnumerationDictionary(lpEnumType As Type) As IDictionary(Of String, Integer)
      Try
        Dim lobjDictionary As New SortedDictionary(Of String, Integer)


        Dim lstrNames As IEnumerable(Of String)
        'Dim lintValues As IEnumerable(Of Integer)

        If Not lpEnumType.IsEnum Then
          Throw New ArgumentOutOfRangeException(NameOf(lpEnumType),
          String.Format("The type '{0}' is not an enumeration.", lpEnumType.Name))
        End If

        'lintValues = [Enum].GetValues(lpEnumType)
        lstrNames = [Enum].GetNames(lpEnumType)

        'For Each lintValue As Object In lintValues
        '  lobjDictionary.Add([Enum].GetName(lpEnumType, lintValue), lintValue)
        'Next

        For Each lstrName As String In lstrNames
          lobjDictionary.Add(lstrName, [Enum].Parse(lpEnumType, lstrName))
        Next

        Return lobjDictionary

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Evaluates the candidate string to determine whether or not it is a valid Guid.
    ''' </summary>
    ''' <param name="lpCandidate">The candidate Guid string.</param>
    ''' <param name="lpOutput">A Guid object reference that is created if the Guid candidate is valid.</param>
    ''' <returns>True if the candidate is a valid Guid, otherwise False.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsGuid(ByVal lpCandidate As String, ByRef lpOutput As Guid) As Boolean

      Try

        Dim lobjIsGuid As New Regex("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$", RegexOptions.Compiled)

        Dim lblnIsValid As Boolean = False
        lpOutput = Guid.Empty

        If Not String.IsNullOrEmpty(lpCandidate) Then
          If lobjIsGuid.IsMatch(lpCandidate) Then
            lpOutput = New Guid(lpCandidate)
            lblnIsValid = True
          End If
        End If

        Return lblnIsValid

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Shared Sub VerifyFilePath(ByVal lpFilePath As String,
                               ByVal lpShouldExist As Boolean)
      Try

        VerifyFilePath(lpFilePath, "lpFilePath", lpShouldExist)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub VerifyFilePath(ByVal lpFilePath As String,
                                   ByVal lpFileParameterName As String,
                                   ByVal lpShouldExist As Boolean)
      Try

        ' Check to see if we have a valid archivePath


        If String.IsNullOrEmpty(lpFilePath) Then
          Throw New ArgumentNullException(lpFileParameterName, "Please specify a value for the file path.")
        ElseIf lpShouldExist = True AndAlso IO.File.Exists(lpFilePath) = False Then
          Dim lobjFileEx As New FileNotFoundException(
          String.Format("The path '{0}' does not point to a valid file.", lpFilePath))
          Throw lobjFileEx
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Checks the specified stream to see if it represents a zip file.
    ''' </summary>
    ''' <param name="lpStream">The stream to check.</param>
    ''' <returns>True if the first line starts with 
    ''' a zip header, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsZipStream(ByVal lpStream As Stream) As Boolean
      Try

        If (lpStream.CanSeek = False) Then
          Throw New Exception("IsZipStream requires a stream that is seekable.")
        End If

        'Using lobjStreamReader As New StreamReader(lpStream)
        Dim lobjStreamReader As New StreamReader(lpStream)
        lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
        Dim lobjBuffer As Char()
        ReDim lobjBuffer(1)
        lobjStreamReader.ReadBlock(lobjBuffer, 0, 2)
        Dim lstrFirstLine As New String(lobjBuffer)
        lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
        If String.Equals("PK", lstrFirstLine, StringComparison.InvariantCultureIgnoreCase) Then
          Return True
        Else
          Return False
        End If
        'End Using
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Reads data from a stream until the end is reached. The
    ''' data is returned as a byte array. An IOException is
    ''' thrown if any of the underlying IO calls fail.
    ''' </summary>
    ''' <param name="lpStream">The stream to read data from</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ReadStreamToByteArray(lpStream As Stream) As Byte()
      Try
        Return ReadStreamToByteArray(lpStream, lpStream.Length)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Reads data from a stream until the end is reached. The
    ''' data is returned as a byte array. An IOException is
    ''' thrown if any of the underlying IO calls fail.
    ''' </summary>
    ''' <param name="lpStream">The stream to read data from</param>
    ''' <param name="lpInitialLength">The initial buffer length</param>
    Public Shared Function ReadStreamToByteArray(lpStream As Stream, lpInitialLength As Integer) As Byte()
      ' If we've been passed an unhelpful initial length, just
      ' use 32K.
      Try
        If lpInitialLength < 1 Then
          lpInitialLength = 32768
        End If

        Dim lobjBuffer As Byte() = New Byte(lpInitialLength - 1) {}
        Dim lintRead As Integer = 0

        Dim lintChunk As Integer

        ' While (lintChunk = lpStream.Read(lobjBuffer, lintRead, lobjBuffer.Length - lintRead)) > 0
        lintChunk = lpStream.Read(lobjBuffer, lintRead, lobjBuffer.Length - lintRead)
        Do While lintChunk > 0
          lintRead += lintChunk

          ' If we've reached the end of our buffer, check to see if there's
          ' any more information
          If lintRead = lobjBuffer.Length Then
            Dim lintNextByte As Integer = lpStream.ReadByte()

            ' End of stream? If so, we're done
            If lintNextByte = -1 Then
              Return lobjBuffer
            End If

            ' Nope. Resize the buffer, put in the byte we've just
            ' read, and continue
            Dim lobjNewBuffer As Byte() = New Byte(lobjBuffer.Length * 2 - 1) {}
            Array.Copy(lobjBuffer, lobjNewBuffer, lobjBuffer.Length)
            lobjNewBuffer(lintRead) = CByte(lintNextByte)
            lobjBuffer = lobjNewBuffer
            lintRead += 1
          End If
          lintChunk = lpStream.Read(lobjBuffer, lintRead, lobjBuffer.Length - lintRead)
        Loop
        ' Buffer is now too big. Shrink it.
        Dim lobjReturnArray As Byte() = New Byte(lintRead - 1) {}
        Array.Copy(lobjBuffer, lobjReturnArray, lintRead)
        Return lobjReturnArray
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''   Checks for folder paths with redundant backslashes and 
    '''   UNC paths with only a single leading backslash.
    ''' </summary>
    ''' <param name="lpSourcePath" type="String">
    '''     <para>
    '''         The path to check.
    '''     </para>
    ''' </param>
    ''' <remarks>
    '''   This is to clean up paths with too many backslashes or incorrect UNC 
    '''   paths which are often the product of a well intentioned effort to 
    '''   eliminate duplicate backslashes in the middle of a path.
    ''' </remarks>
    ''' <returns>
    '''   If a path with a single leading backslash was provided a clean path 
    '''   with a leading double backslash will be returned, otherwise the 
    '''   cleaned original path will be returned.
    ''' </returns>
    Public Shared Function CleanPath(lpSourcePath As String) As String
      Try
        Dim lobjCleanPathBuilder As New StringBuilder(lpSourcePath.Replace("\\", "\"))
        If lobjCleanPathBuilder.Chars(0) = "\" AndAlso lobjCleanPathBuilder.Chars(1) <> "\" Then
          lobjCleanPathBuilder.Insert(0, "\")
        End If

        Return lobjCleanPathBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' GetProtocol
    ''' </summary>
    ''' <param name="lpPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetProtocol(lpPath As String) As Network.ProtocolEnum

      Try
        Dim url As New Uri(lpPath)
        Dim lstrScheme As String = url.Scheme

        If Not lstrScheme.Contains("."c) Then
#If SILVERLIGHT = 1 Then
        Return [Enum].Parse(GetType(Network.ProtocolEnum), lstrScheme, True)
#Else
          Return [Enum].Parse(GetType(Network.ProtocolEnum), lstrScheme)
#End If

        ElseIf lstrScheme.Equals("net.pipe") Then
          Return Network.ProtocolEnum.pipe
        ElseIf lstrScheme.Equals("net.tcp") Then
          Return Network.ProtocolEnum.tcp
        Else
          Return -1
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Shared Function CopyStream(ByVal lpOriginalStream As Stream) As MemoryStream
      Try
        Dim lintChunkSize As Integer = 1024
        Dim lobjBuffer(lintChunkSize - 1) As Byte
        Dim numBytesRead As Integer = 0
        Dim lobjOutputStream As New IO.MemoryStream()

        With lpOriginalStream
          If (.CanSeek) Then
            .Position = 0
            lobjOutputStream.SetLength(.Length)
          End If
          numBytesRead = .Read(lobjBuffer, 0, lintChunkSize)

          While (numBytesRead > 0)
            lobjOutputStream.Write(lobjBuffer, 0, numBytesRead)
            numBytesRead = .Read(lobjBuffer, 0, lintChunkSize)
          End While
        End With

        If lpOriginalStream.CanSeek Then
          lpOriginalStream.Position = 0
        End If

        lobjOutputStream.Position = 0

        ' Clear the byte array
        lobjBuffer = Nothing

        Return lobjOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub CopyStream(ByVal lpOriginalStream As Stream, lpDestinationStream As Stream)
      Try
        Dim lintChunkSize As Integer = 1024
        Dim lobjBuffer(lintChunkSize - 1) As Byte
        Dim numBytesRead As Integer = 0
        'Dim lobjOutputStream As New IO.MemoryStream()

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpDestinationStream)
#Else
          If lpDestinationStream Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpDestinationStream))
          End If
#End If

        If Not lpDestinationStream.CanWrite Then
          Throw New ArgumentOutOfRangeException(NameOf(lpDestinationStream), "Destination stream is not writable.")
        End If

        With lpOriginalStream
          If (.CanSeek) Then
            .Position = 0
            If lpDestinationStream.CanSeek Then
              lpDestinationStream.SetLength(.Length)
            End If
          End If
          numBytesRead = .Read(lobjBuffer, 0, lintChunkSize)

          While (numBytesRead > 0)
            lpDestinationStream.Write(lobjBuffer, 0, numBytesRead)
            numBytesRead = .Read(lobjBuffer, 0, lintChunkSize)
          End While
        End With

        If lpOriginalStream.CanSeek Then
          lpOriginalStream.Position = 0
        End If

        If lpDestinationStream.CanSeek Then
          lpDestinationStream.Position = 0
        End If

        ' Clear the byte array
        lobjBuffer = Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function CopyStreamToString(ByVal lpStream As Stream) As String
      Try
        Return CopyStreamToString(lpStream, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CopyStreamToString(ByVal lpStream As Stream, ByVal lpEncoding As Encoding) As String
      Try
        Return CopyStreamToString(lpStream, True, lpEncoding)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CopyStreamToString(ByVal lpStream As Stream, ByVal lpStartAtBeginning As Boolean) As String

      Try

        Return CopyStreamToString(lpStream, lpStartAtBeginning, Encoding.Default)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CopyStreamToString(ByVal lpStream As Stream, ByVal lpStartAtBeginning As Boolean, ByVal lpEncoding As Encoding) As String

      Dim lstrReturnString As String

      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpStream)
#Else
          If lpStream Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpStream))
          End If
#End If

        If lpStream.CanRead = False Then
          Throw New InvalidOperationException("Can't read stream, it is not readable.")
        End If

        If lpStartAtBeginning = True AndAlso lpStream.CanSeek Then
          lpStream.Position = 0
        End If

        Dim lobjStreamReader As New StreamReader(lpStream, lpEncoding)
        'lobjStreamReader.CurrentEncoding=Encoding.
        lstrReturnString = lobjStreamReader.ReadToEnd

        If lpStartAtBeginning = True AndAlso lpStream.CanSeek Then
          lpStream.Position = 0
        End If

        Return lstrReturnString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Writes the specified file to a memory stream.
    ''' </summary>
    ''' <param name="lpFilePath">The file to be read.</param>
    ''' <returns>A MemoryStream containing the contents of the specified file.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If the lpFilePath argument is null, an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="FileNotFoundException">
    ''' If the specified file does not exist, a FileNetFoundException will be thrown.
    ''' </exception>
    Public Shared Function WriteFileToMemoryStream(ByVal lpFilePath As String) As MemoryStream
      Dim numBytesRead As Integer = 0
      Try
        ' Stream the file
        'Dim lobjFileStream As FileStream
        'Dim lobjMemoryStream As MemoryStream

        Dim lenuProtocol As Network.ProtocolEnum = GetProtocol(lpFilePath)

        Select Case lenuProtocol
          Case Network.ProtocolEnum.http, Network.ProtocolEnum.https

            Dim lobjWebClient As New WebClient With {
              .Credentials = CredentialCache.DefaultCredentials
            }

            Dim lobjBytes As Byte() = lobjWebClient.DownloadData(lpFilePath)
            Return New MemoryStream(lobjBytes)

          Case Else
            VerifyFilePath(lpFilePath, True)

            If File.Exists(lpFilePath) = False Then
              Throw New FileNotFoundException(String.Format("Unable to write file to stream, the file '{0}' could not be found.", lpFilePath)) ', lpFilePath)
            End If

            Dim lintChunkSize As Integer = 1024
            Dim lobjBuffer(lintChunkSize - 1) As Byte

            Dim lobjMemoryStream As New IO.MemoryStream()

            Using lobjFileStream As FileStream = File.OpenRead(lpFilePath)
              lobjMemoryStream.SetLength(lobjFileStream.Length)
              numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)

              While (numBytesRead > 0)
                lobjMemoryStream.Write(lobjBuffer, 0, numBytesRead)
                numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)
              End While
            End Using

            lobjMemoryStream.Position = 0

            Return lobjMemoryStream

        End Select



      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ShortenPath(lpSourcePath As String, Optional lpMaxPathLength As Integer = 260) As String
      Try

        If lpSourcePath.Length <= lpMaxPathLength Then
          Return lpSourcePath
        End If

        Dim lstrFolderPath As String = Path.GetDirectoryName(lpSourcePath)
        If lstrFolderPath.Length > lpMaxPathLength Then
          Throw New InvalidOperationException("The folder path is too long to shorten.")
        End If

        Dim lstrFileNameOnly As String = Path.GetFileNameWithoutExtension(lpSourcePath)
        Dim lstrExtension As String = Path.GetExtension(lpSourcePath)

        Dim lintOriginalFileNameLength As Integer = lpSourcePath.Length
        Dim lintFolderPathLength As Integer = lstrFolderPath.Length
        Dim lintFileNameOnlyLength As Integer = lstrFileNameOnly.Length
        Dim lintExtensionLength As Integer = lstrExtension.Length

        Dim lintNewFileNameLength As Integer = lpMaxPathLength - lintFolderPathLength - lintExtensionLength - 1

        If lintNewFileNameLength < 1 Then
          Throw New InvalidOperationException("The file name is too long to shorten.")
        End If

        Dim lstrNewFileName As String = String.Concat(lstrFileNameOnly.AsSpan(0, lintNewFileNameLength), lstrExtension)

        Return Path.Combine(lstrFolderPath, lstrNewFileName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CopyStreamToByteArray(ByVal lpOriginalStream As Stream) As Byte()
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpOriginalStream)
#Else
          If lpOriginalStream Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpOriginalStream))
          End If
#End If

        If lpOriginalStream.CanRead = False Then
          Throw New InvalidOperationException("Stream can't be read.")
        End If

        If TypeOf lpOriginalStream Is MemoryStream Then
          lpOriginalStream.Position = 0
          Return CType(lpOriginalStream, MemoryStream).ToArray
        Else

          Return CopyStream(lpOriginalStream).ToArray()

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the specified string is a number.
    ''' </summary>
    ''' <param name="lpValue">The string to check.</param>
    ''' <returns></returns>
    ''' <remarks>Similiar in use to Visual Basic IsNumeric.  
    ''' First tries to parse as a whole number, 
    ''' then tries to parse as a floating number.</remarks>
    Public Shared Function IsNumeric(ByVal lpValue As String) As Boolean
      Try

        Dim lblnReturnValue As Boolean
        Dim llngResult As Long
        Dim ldblResult As Double

        ' Try it as a long first
        lblnReturnValue = Long.TryParse(lpValue, llngResult)

        If lblnReturnValue = False Then
          ' Then try it as a double
          lblnReturnValue = Double.TryParse(lpValue, ldblResult)
        End If

        ' Return the result
        Return lblnReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CopyStringToStream(lpContent As String) As Stream
      Try

        Dim lobjOutputStream As New MemoryStream
        Dim lobjStreamWriter As New StreamWriter(lobjOutputStream)
        lobjStreamWriter.Write(lpContent)
        lobjStreamWriter.Flush()

        lobjOutputStream.Position = 0

        Return lobjOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     Gets all the text from a text file.
    ''' </summary>
    ''' <param name="lpSourcePath" type="String">
    '''     <para>
    '''         The fully qualified path of the source file.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     The complete text of the file.
    ''' </returns>
    Public Shared Function ReadAllTextFromFile(lpSourcePath As String) As String
      Try

        Dim lstrReturnString As String = Nothing

        Using lobjStreamReader As New StreamReader(lpSourcePath)
          lstrReturnString = lobjStreamReader.ReadToEnd
        End Using

        Return lstrReturnString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ByteArrayToString(lpInput As Byte()) As String
      Try

        Dim lobjBuilder As New StringBuilder(lpInput.Length)
        For lintByteCounter As Integer = 0 To lpInput.Length - 1
          lobjBuilder.Append(lpInput(lintByteCounter).ToString("X2"))
        Next

        Return lobjBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Evaluates the candidate string to determine whether or not it contains a valid Guid.
    ''' </summary>
    ''' <param name="lpCandidate">The candidate Guid string.</param>
    ''' <param name="lpOutput">A Guid object reference that is created if the Guid candidate is valid.</param>
    ''' <returns>True if the candidate contains a valid Guid, otherwise False.</returns>
    ''' <remarks>This method is used in cases where a string is not a guid but might contain a guid.</remarks>
    Public Shared Function HasGuid(ByVal lpCandidate As String, ByRef lpOutput As Guid) As Boolean
      Try

        Dim lobjIsGuid As New Regex("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$", RegexOptions.Compiled)

        Dim lblnHasGuid As Boolean = False
        Dim lobjMatchCollection As MatchCollection = Nothing

        lpOutput = Guid.Empty

        If Not String.IsNullOrEmpty(lpCandidate) Then
          lobjMatchCollection = lobjIsGuid.Matches(lpCandidate)
          If lobjMatchCollection.Count > 0 Then
            lpOutput = New Guid(lobjMatchCollection.Item(0).Value)
            lblnHasGuid = True
          End If
        End If

        Return lblnHasGuid

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

  End Class

End Namespace