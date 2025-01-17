Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Text
Imports System.Xml
Imports Microsoft.Data.SqlClient
Namespace Utilities

  Partial Public Class Helper

#Region "General Helper Methods"

    Public Shared Function ArrayContainsValue(ByVal lpArray As Array, ByVal lpValue As String) As Boolean

      Try

        For lintElementCounter As Int16 = 0 To lpArray.Length - 1
          If lpArray(lintElementCounter) = lpValue Then
            Return True
          End If
        Next

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False

      End Try
    End Function

#End Region

#Region "XML Methods"

    ''' <summary>
    ''' Returns formatted xml string (indent and newlines) from unformatted XML
    ''' string for display in eg textboxes.
    ''' </summary>
    ''' <param name="lpUnformattedXml">Unformatted xml string.</param>
    ''' <returns>Formatted xml string and any exceptions that occur.</returns>
    Public Shared Function FormatXmlString(ByVal lpUnformattedXml As String) As String
      Try
        ' Load unformatted xml into a dom
        Dim lobjXmlDocument As New XmlDocument()
        lobjXmlDocument.LoadXml(lpUnformattedXml)

        ' Will hold formatted xml
        Dim lobjStringBuilder As New StringBuilder

        ' Pumps the formatted xml into the StringBuilder above
        Dim lobjStringWriter As New StringWriter(lobjStringBuilder)

        ' Does the formatting
        Dim lobjXmlTextWriter As XmlTextWriter = Nothing

        Try
          ' Point the XmlTextWriter at the StringWriter
          ' We want the output formatted
          lobjXmlTextWriter = New XmlTextWriter(lobjStringWriter) With {
            .Formatting = Formatting.Indented
          }

          ' Get the dom to dump its contents into the XmlTextWriter
          lobjXmlDocument.WriteTo(lobjXmlTextWriter)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Finally
          ' Clean up even if we get an error
          lobjXmlTextWriter?.Close()
        End Try

        ' return the formatted xml
        Return lobjStringBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Formats the specified xml file (i.e. indent and newlines)
    ''' </summary>
    ''' <param name="lpXmlFilePath">A fully qualified path to the xml file</param>
    ''' <remarks></remarks>
    Public Shared Sub FormatXmlFile(ByVal lpXmlFilePath As String)
      Try
        Dim lobjXmlDocument As New XmlDocument

        lobjXmlDocument.Load(lpXmlFilePath)

        lobjXmlDocument = FormatXmlDocument(lobjXmlDocument)

        lobjXmlDocument.Save(lpXmlFilePath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Returns formatted xml document (indent and newlines) from unformatted XML
    ''' document for display in eg textboxes.
    ''' </summary>
    ''' <param name="lpXmlDocument">Unformatted xml document.</param>
    ''' <returns>Formatted xml string and any exceptions that occur.</returns>
    Public Shared Function FormatXmlDocument(ByVal lpXmlDocument As XmlDocument) As XmlDocument
      Try


        ' Get the xml string from the original document
        Dim lstrUnformatedXml As String = lpXmlDocument.OuterXml

        ' Format the string
        Dim lstrFormatedXml As String = FormatXmlString(lstrUnformatedXml)

        ' Load the formatted xml string into a new document
        Dim lobjXmlDocument As New XmlDocument()
        lobjXmlDocument.LoadXml(lstrFormatedXml)

        ' Return the new document
        Return lobjXmlDocument

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Public Shared Function RemoveEntriesFromString(lpOriginalString As String,
                                                   ByVal ParamArray lpDuplicateEntryCandidates() As String) As String
      Try
        Return RemoveEntriesFromString(lpOriginalString, 0, lpDuplicateEntryCandidates)
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function RemoveEntriesFromString(lpOriginalString As String,
                                                   lpStartPosition As Integer,
                                                   ByVal ParamArray lpDuplicateEntryCandidates() As String) As String
      Try

        Dim lstrReturnString As String = lpOriginalString
        Dim lintEntryIndex As Integer

        For Each lstrDuplicateEntryCandidate As String In lpDuplicateEntryCandidates
          ' lstrReturnString = lstrReturnString.Substring(lpStartPosition).Replace(lstrDuplicateEntryCandidate, String.Empty)
          lintEntryIndex = lstrReturnString.IndexOf(lstrDuplicateEntryCandidate, lpStartPosition)
          If lintEntryIndex > lpStartPosition Then
            lstrReturnString = lstrReturnString.Remove(lintEntryIndex, lstrDuplicateEntryCandidate.Length)
          End If
        Next

        Return lstrReturnString

      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetResourceFileFromAssembly(lpResourceFileName As String, lpAssembly As Assembly) As Stream
      Try

        Dim lobjResourceStream As Stream = Nothing
        Dim lstrAssemblyFullName As String = lpAssembly.FullName
        'Dim lstrResourceAssemblyName As String = String.Format("{0}.{1}",
        '                                                       lstrAssemblyFullName.Substring(0,
        '                                                                                      lstrAssemblyFullName.
        '                                                                                       IndexOf(","c)),
        '                                                       lpResourceFileName)

        'lobjResourceStream = lpAssembly.GetManifestResourceStream(lstrResourceAssemblyName)

        Dim lstrResourceNames As String() = lpAssembly.GetManifestResourceNames

        If lstrResourceNames.Length = 0 Then
          Throw New FileNotFoundException(
            String.Format("Unable to get resource file '{0}' from assembly '{1}', assembly has no embedded resource files.",
                          lpResourceFileName, lstrAssemblyFullName), lpResourceFileName)
        End If


#If NET8_0_OR_GREATER Then
        Dim lstrResourceName = (From item In lstrResourceNames Where item.Contains(lpResourceFileName, StringComparison.CurrentCultureIgnoreCase) Select item).Single
#Else
        Dim lstrResourceName = (From item In lstrResourceNames Where item.Contains(lpResourceFileName) Select item).Single
#End If

        If Not String.IsNullOrEmpty(lstrResourceName) Then
          lobjResourceStream = lpAssembly.GetManifestResourceStream(lstrResourceName)
        End If

        If lobjResourceStream Is Nothing Then
          Throw New FileNotFoundException(
            String.Format("Unable to get resource file '{0}' from assembly '{1}', file is not an embedded resource.",
                          lpResourceFileName, lstrAssemblyFullName), lpResourceFileName)
        End If

        If lobjResourceStream.CanSeek Then
          lobjResourceStream.Position = 0
        End If

        Return lobjResourceStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetResourceFileTextFromAssembly(lpResourceFileName As String, lpAssembly As Assembly) _
      As String
      Try
        Dim lstrResourceString As String = Nothing
        Dim lobjResourceFileStream As Stream = Helper.GetResourceFileFromAssembly(lpResourceFileName, lpAssembly)

        Using lobjStreamReader As New StreamReader(lobjResourceFileStream)
          lstrResourceString = lobjStreamReader.ReadToEnd
          lobjStreamReader.Close()
        End Using

        Return lstrResourceString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetValueFromDataRecord(ByVal lpDataRecord As IDataRecord,
                                                  ByVal lpColumnName As String,
                                                  ByVal lpDefaultValue As Object) As Object
      Try
        If HasColumn(lpDataRecord, lpColumnName) Then
          If IsDBNull(lpDataRecord(lpColumnName)) Then
            Return lpDefaultValue
          Else
            Return lpDataRecord(lpColumnName)
          End If
        Else
          Return lpDefaultValue
        End If
      Catch StackEx As StackOverflowException
        ApplicationLogging.LogException(StackEx, Reflection.MethodBase.GetCurrentMethod)
        Return lpDefaultValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub HandleConnection(ByVal lpConnection As SqlConnection)
      Try
        With lpConnection
          'do a switch on the state of the connection
          Select Case .State
            Case ConnectionState.Open
              'RKS - Did some testing and by closing and re-opening causes a MAJOR performance impact.
              'If it's already open, why do we need to close and re-open?
              'the connection is open
              'close then re-open
              '.Close() 
              '.Open()
              Exit Select
            Case ConnectionState.Closed
              'connection is not open
              'open the connection
              .Open()
              Exit Select
            Case Else
              .Close()
              .Open()
              Exit Select
          End Select
        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function HasColumn(lpDataRecord As IDataRecord, lpColumnName As String) As Boolean
      Try
        For lintColumnCounter As Integer = 0 To lpDataRecord.FieldCount - 1
          If lpDataRecord.GetName(lintColumnCounter).Equals(lpColumnName, StringComparison.InvariantCultureIgnoreCase) _
            Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the file is locked by another process.
    ''' </summary>
    ''' <param name="lpFilePath">The path of the file to check.</param>
    ''' <returns>True if the file is locked, otherwise false.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="FileNotFoundException">
    ''' If a bad file path is provided a FileNotFoundException will be thrown.
    ''' </exception>
    Public Shared Function IsFileLocked(lpFilePath As String) As Boolean
      Try

        If File.Exists(lpFilePath) = False Then
          Throw New FileNotFoundException
        End If

        Using lobjFileStream As FileStream = File.OpenRead(lpFilePath)
          lobjFileStream.Close()
          Return False
        End Using

      Catch BadPathEx As FileNotFoundException
        ApplicationLogging.LogException(BadPathEx, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Catch IoEx As IOException
        Return True
      Catch UnAuthEx As UnauthorizedAccessException
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return True
      End Try
    End Function

    ''' <summary>Checks to see if the URL is valid and available.</summary>
    ''' <param name="lpUrl">The url to check.</param>
    ''' <exception caption="" cref="Exceptions.ServerUnavailableException">
    ''' If the URL does not return a valid response code a ServerUnavailableException will be thrown.
    ''' </exception>
    Public Shared Function IsUrlAvailable(ByVal lpUrl As String) As Boolean
      Try
        Return IsUrlAvailable(lpUrl, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Checks to see if the URL is valid and available.</summary>
    ''' <param name="lpUrl">The url to check.</param>
    ''' <param name="lpStatusCode">Gets the string representation of the returned http status code.</param>
    Public Shared Function IsUrlAvailable(ByVal lpUrl As String, ByRef lpStatusCode As String) As Boolean
      Try
        Dim lobjRequest As HttpWebRequest
        Dim lobjWebResponse As HttpWebResponse

        Try
          lobjRequest = WebRequest.Create(lpUrl)
          lobjWebResponse = lobjRequest.GetResponse()
        Catch WebEx As WebException
          ' Throw New Exceptions.ServerUnavailableException(WebEx.Message, lpUrl, WebEx)
          Return False
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

        lpStatusCode = lobjWebResponse.StatusCode.ToString
        If lobjWebResponse.StatusCode = HttpStatusCode.OK Then
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
    ''' This Function will be used to check if the passed String argument is null.
    ''' If it is null it will return the empty string, otherwise it returns the passed string
    ''' </summary>
    ''' <param name="lpValue">The string value to check</param>
    ''' <returns>
    ''' If the passed value is null it returns an empty string.  
    ''' If not it returns the passed string.
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function CheckForNullString(ByVal lpValue As String) As String

      Try

        If lpValue Is Nothing Then
          Return String.Empty
        Else
          Return lpValue
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function IsAssignableFrom(lpInterface As String, lpObject As Object) As Boolean
      Try
        'Return lpInterface.IsAssignableFrom(lpObject.GetType)
        Dim lobjInterface As Type = lpObject.GetType.GetInterface(lpInterface)
        If lobjInterface IsNot Nothing Then
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

    Public Shared Function IsAssignableFrom(lpInterface As String, lpType As Type) As Boolean
      Try
        'Return lpInterface.IsAssignableFrom(lpObject.GetType)
        Dim lobjInterface As Type = lpType.GetInterface(lpInterface)
        If lobjInterface IsNot Nothing Then
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
    ''' Determines whether or not we are running a deployed instance of the application or not.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>This is based on the assumption that if the application path is in a 'bin' 
    ''' or 'shared' directory that we are not deployed but are in the development environment.</remarks>
    Public Shared Function IsRunningInstalled() As Boolean
      Try

        Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location) 'My.Application.Info.DirectoryPath
        Dim lstrConfigPath As String = ""
        If assemblyPath.Contains("bin") OrElse assemblyPath.ToLower.Contains("shared") Then
          Return False
        Else
          Return True
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function IsPdfStream(ByVal lpStream As Stream) As Boolean
    '  Try
    '    If PdfSharp.Pdf.IO.PdfReader.TestPdfFile(lpStream) <> 0 Then
    '      Return True
    '    Else
    '      Return False
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ' <Added by: Ernie at: 9/24/2013-3:17:35 PM on machine: ERNIE-THINK>
    ' The tiff test method below was ported from a code listing on a stackoverflow article
    ' http://stackoverflow.com/questions/2731917/how-to-detect-if-a-file-is-pdf-or-tiff
    ' It was posted by none other than Steve Hawley (Chief Architect at Atalasoft)
    ' http://www.atalasoft.com/blogs/stevehawley
    ' They are the same vendor we previously purchased an imaging library from.
    Public Shared Function IsTiffStream(ByVal lpStream As Stream) As Boolean
      Try

        ' Const TIFF_TAG_LENGTH As Integer = 12
        Const HEADER_SIZE As Integer = 2
        Const MINIMUM_TIFF_SIZE As Integer = 8
        Const INTEL_MARK As Byte = &H49
        Const MOTOROLA_MARK As Byte = &H4D
        Const TIFF_MAGIC_NUMBER As UShort = 42

        lpStream.Seek(0, SeekOrigin.Begin)
        If lpStream.Length < MINIMUM_TIFF_SIZE Then
          Return False
        End If
        Dim lobjHeader As Byte() = New Byte(HEADER_SIZE - 1) {}

        lpStream.Read(lobjHeader, 0, lobjHeader.Length)

        If lobjHeader(0) <> lobjHeader(1) OrElse (lobjHeader(0) <> INTEL_MARK AndAlso lobjHeader(0) <> MOTOROLA_MARK) _
          Then
          lpStream.Seek(0, SeekOrigin.Begin)
          Return False
        End If
        Dim lblnIsIntel As Boolean = lobjHeader(0) = INTEL_MARK

        Dim lintMagicNumber As UShort = ReadShort(lpStream, lblnIsIntel)

        lpStream.Seek(0, SeekOrigin.Begin)

        If lintMagicNumber <> TIFF_MAGIC_NUMBER Then
          Return False
        End If
        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function IsTiffFile(ByVal lpFilePath As String) As Boolean
      Try
        Dim lblnIsTiff As Boolean

        Using lobjFileStream As New FileStream(lpFilePath, FileMode.Open, FileAccess.Read)
          lblnIsTiff = IsTiffStream(lobjFileStream)
          lobjFileStream.Close()
          lobjFileStream.Dispose()
        End Using

        Return lblnIsTiff

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function


#Region "Statistical Methods"

    Public Shared Function Total(lpValues As IEnumerable(Of Double), Optional ByRef lpSize As Integer = 0) As Double
      Try
        Dim lintSize As Integer
        Dim ldblSum As Double = 0

        For Each lblnValue As Double In lpValues
          lintSize += 1
          ldblSum += lblnValue
        Next

        lpSize = lintSize
        Return ldblSum

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Average(lpValues As IEnumerable(Of Double), Optional ByRef lpSize As Integer = 0) As Double
      Try

        Dim lintSize As Integer
        Dim ldblSum As Double = Total(lpValues, lintSize)

        lpSize = lintSize

        Return ldblSum / lintSize

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function StDev(lpValues As IEnumerable(Of Double)) As Double
      Try
        Dim ldblAverage As Double = lpValues.Average()
        Return Math.Sqrt(lpValues.Average(Function(v) Math.Pow(v - ldblAverage, 2)))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' The methods below are based on a codeproject article
    ' http://www.codeproject.com/Articles/42492/Using-LINQ-to-Calculate-Basic-Statistics

    Public Shared Function Variance(ByVal lpSource As IEnumerable(Of Double)) As Double
      Try
        Dim n As Integer = 0
        Dim mean As Double = 0
        Dim M2 As Double = 0

        For Each x As Double In lpSource
          n += 1
          Dim delta As Double = x - mean
          mean += delta / n
          M2 += delta * (x - mean)
        Next x
        Return M2 / (n - 1)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Variance(ByVal lpSource As IEnumerable(Of FileSize)) As Double
      Try
        Dim n As Integer = 0
        Dim mean As Double = 0
        Dim M2 As Double = 0

        For Each x As FileSize In lpSource
          n += 1
          Dim delta As Double = x.Bytes - mean
          mean += delta / n
          M2 += delta * (x.Bytes - mean)
        Next x
        Return M2 / (n - 1)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function StandardDeviation(ByVal lpSource As IEnumerable(Of Double)) As Double
      Try
        Return Math.Sqrt(Variance(lpSource))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Median(ByVal source As IEnumerable(Of Double)) As Double
      Try
        'Dim sortedList = From number In source
        '      Order By number
        '      Select number

        Dim sortedList As New List(Of Double)

        sortedList.AddRange(From number In source
                            Order By number
                            Select number)

        Dim count As Integer = sortedList.Count()
        Dim itemIndex As Integer = count \ 2
        If count Mod 2 = 0 Then ' Even number of items.
          Return (sortedList.ElementAt(itemIndex) + sortedList.ElementAt(itemIndex - 1)) / 2
        End If

        ' Odd number of items. 
        Return sortedList.ElementAt(itemIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function Median (Of T As Structure)(ByVal source As IEnumerable(Of T)) As T?
    '  Try
    '    'Dim sortedList = From number In source
    '    '                 Order By number
    '    '                 Select number




    '    Dim sortedList As New List(Of T)

    '    sortedlist.AddRange( From number In source
    '          Order By number
    '          Select number)


    '    Dim count As Integer = sortedList.Count()
    '    Dim itemIndex As Integer = count \ 2
    '    If count Mod 2 = 0 Then ' Even number of items.
    '      Return (DirectCast(sortedList.Item(itemIndex), Object) + DirectCast(sortedList.Item(itemIndex - 1), Object)) / 2

    '    End If

    '    ' Odd number of items. 
    '    Return sortedList.Item(itemIndex)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Shared Function Median(ByVal source As IEnumerable(Of FileSize)) As FileSize
      Try
        'Dim sortedList = From number In source
        '                 Order By number
        '                 Select number




        Dim sortedList As New List(Of FileSize)

        sortedList.AddRange(From number In source
                            Order By number
                            Select number)


        Dim count As Integer = sortedList.Count()
        Dim itemIndex As Integer = count \ 2
        If count Mod 2 = 0 Then ' Even number of items.
          Return New FileSize(CLng((sortedList.Item(itemIndex).Bytes + sortedList.Item(itemIndex - 1).Bytes) / 2))

        End If

        ' Odd number of items. 
        Return sortedList.Item(itemIndex)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Median(ByVal source As IEnumerable(Of TimeSpan)) As TimeSpan
      Try
        'Dim sortedList = From number In source
        '                 Order By number
        '                 Select number




        Dim sortedList As New List(Of TimeSpan)

        sortedList.AddRange(From number In source
                            Order By number
                            Select number)


        Dim count As Integer = sortedList.Count()
        Dim itemIndex As Integer = count \ 2
        If count Mod 2 = 0 Then ' Even number of items.
          Return New TimeSpan((sortedList.Item(itemIndex).Ticks + sortedList.Item(itemIndex - 1).Ticks) / 2)

        End If

        ' Odd number of items. 
        Return sortedList.Item(itemIndex)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Mode(Of T As Structure)(ByVal source As IEnumerable(Of T)) As T?
      Try
        Dim sortedList = From number In source
                         Order By number
                         Select number

        Dim count As Integer = 0
        Dim max As Integer = 0
        Dim current As T = Nothing
        Dim mode_Renamed As New T?()

        For Each [next] As T In sortedList
          If current.Equals([next]) = False Then
            current = [next]
            count = 1
          Else
            count += 1
          End If

          If count > max Then
            max = count
            mode_Renamed = current
          End If
        Next [next]

        If max > 1 Then
          Return mode_Renamed
        End If

        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Range(ByVal source As IEnumerable(Of Double)) As Double
      Try
        Return source.Max() - source.Min()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Range(ByVal source As IEnumerable(Of FileSize)) As FileSize
      Try
        Return source.Max() - source.Min()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Range(ByVal source As IEnumerable(Of TimeSpan)) As TimeSpan
      Try
        Return source.Max() - source.Min()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Covariance(ByVal source As IEnumerable(Of Double), ByVal other As IEnumerable(Of Double)) _
      As Double
      Try
        Dim len As Integer = source.Count()

        Dim avgSource As Double = source.Average()
        Dim avgOther As Double = other.Average()
        Dim covariance_Renamed As Double = 0

        For i As Integer = 0 To len - 1
          covariance_Renamed += (source.ElementAt(i) - avgSource) * (other.ElementAt(i) - avgOther)
        Next i

        Return covariance_Renamed / len
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "DateTime Methods"

    ''' <summary>
    ''' Renders the specified date time value to a detailed 
    ''' locale specific string including milliseconds using the current culture.
    ''' </summary>
    ''' <param name="lpTimeValue">The date time value to render.</param>
    ''' <returns>A string similiar to '12/5/2011 5:29:31:239 AM'</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDetailedDateString(lpTimeValue As DateTime) As String
      Try
        Return ToDetailedDateString(lpTimeValue, CultureInfo.CurrentCulture.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Renders the specified date time value to a detailed 
    ''' locale specific string including milliseconds.
    ''' </summary>
    ''' <param name="lpTimeValue">The date time value to render.</param>
    ''' <param name="lpCulture">The culture to use when rendering.</param>
    ''' <returns>A string similiar to '12/5/2011 5:29:31:239 AM'</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDetailedDateString(lpTimeValue As DateTime, lpCulture As CultureInfo) As String
      Try

        Dim lobjTimeBuilder As New StringBuilder

        Dim lobjDateTimeFormat As DateTimeFormatInfo = lpCulture.DateTimeFormat

        'lobjDateTimeFormat.
        ' Add the date
        lobjTimeBuilder.AppendFormat("{0} ", lpTimeValue.ToString("d", lobjDateTimeFormat))

        ' Add the time down to the second
        lobjTimeBuilder.AppendFormat("{0}:{1}:{2}", lpTimeValue.Hour, lpTimeValue.Minute.ToString("00"),
                                     lpTimeValue.Second.ToString("00"))

        ' If there are milliseconds, add them also
        If lpTimeValue.Millisecond <> 0 Then
          lobjTimeBuilder.AppendFormat(":{0}", lpTimeValue.Millisecond.ToString("000"))
        End If

        ' Add AM or PM
        If lpTimeValue.Hour < 12 Then
          lobjTimeBuilder.Append(" AM")
        Else
          lobjTimeBuilder.Append(" PM")
        End If

        Return lobjTimeBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Renders the specified date time value to a detailed 
    ''' locale specific string including milliseconds.
    ''' </summary>
    ''' <param name="lpTimeValue">The date time value to render.</param>
    ''' <param name="lpLocale">The locale to use when rendering.</param>
    ''' <returns>A string similiar to '12/5/2011 5:29:31:239 AM'</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDetailedDateString(lpTimeValue As DateTime, lpLocale As String) As String
      Try

        Return ToDetailedDateString(lpTimeValue, New CultureInfo(lpLocale))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a date time value including milliseconds, if available.  
    ''' </summary>
    ''' <param name="lpDateString">The date time string to parse.  
    ''' Expected format '12/5/2011 5:29:31:239 AM'. 
    ''' Format may vary depending on the current culture.</param>
    ''' <returns>A date time value including the milliseconds if available.</returns>
    ''' <remarks>Parses the string based on the current culture.</remarks>
    Public Shared Function FromDetailedDateString(lpDateString As String) As DateTime
      Try
        Return FromDetailedDateString(lpDateString, CultureInfo.CurrentCulture.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a date time value including milliseconds, if available.
    ''' </summary>
    ''' <param name="lpDateString">The date time string to parse.  Expected format '12/5/2011 5:29:31:239 AM'.</param>
    ''' <param name="lpLocale">The locale to use when parsing the date string.</param>
    ''' <returns>A date time value including the milliseconds if available.</returns>
    ''' <remarks></remarks>
    Public Shared Function FromDetailedDateString(lpDateString As String, lpLocale As String) As DateTime
      Try

        If String.IsNullOrEmpty(lpDateString) Then
          Return DateTime.MinValue
        End If

        Dim ldatReturnValue As DateTime = Nothing
        Dim lobjCulture As New CultureInfo(lpLocale)
        Dim lobjDateTimeFormat As DateTimeFormatInfo = lobjCulture.DateTimeFormat
        Dim lstrPrimaryParts As String() = lpDateString.Split(" ")
        Dim lstrTimeParts As String() = lstrPrimaryParts(1).Split(":")
        Dim lintHours As Integer
        Dim lintMinutes As Integer
        Dim lintSeconds As Integer
        Dim lintMilliseconds As Integer

        If String.Equals(lstrPrimaryParts(2), "AM") Then
          lintHours = lstrTimeParts(0)
        Else
          lintHours = lstrTimeParts(0) + 12
        End If

        lintMinutes = lstrTimeParts(1)
        lintSeconds = lstrTimeParts(2)
        lintMilliseconds = lstrTimeParts(3)

        ldatReturnValue = DateTime.Parse(lstrPrimaryParts(0), lobjDateTimeFormat)

        ldatReturnValue = ldatReturnValue.AddHours(lintHours)
        ldatReturnValue = ldatReturnValue.AddMinutes(lintMinutes)
        ldatReturnValue = ldatReturnValue.AddSeconds(lintSeconds)
        ldatReturnValue = ldatReturnValue.AddMilliseconds(lintMilliseconds)

        Return ldatReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Error Handling Methods"

    ''' <summary>
    ''' Loop through the call stack from the bottom up and generate a formatted string.
    ''' </summary>
    ''' <param name="lpException">The exception containing the call stack to format.</param>
    ''' <returns>A formattted string containing the entire call stack.</returns>
    ''' <remarks>Output string only contains line numbers when running in debug mode.</remarks>
    Shared Function FormatCallStack(Optional ByVal lpException As Exception = Nothing) As String

      Try
        ' Create a string to store the call stack in
        ' Dim lstrCallStack As String = ""
        Dim lobjCallStackBuilder As New StringBuilder

        If lpException Is Nothing Then

          ' Create a StackTrace object with file names and line numbers
          Dim lobjStackTrace As New System.Diagnostics.StackTrace(True)

          ' Display the frame information
          For i As Integer = 0 To lobjStackTrace.FrameCount - 1
            With lobjStackTrace.GetFrame(i)
              lobjCallStackBuilder.AppendFormat("Method: {0}, File: '{1}', Line: {2}{3}", .GetMethod.ToString,
                                                  .GetFileName, .GetFileLineNumber, ControlChars.CrLf)
              'lstrCallStack += "Method: " & .GetMethod.ToString & _
              '", File: '" & .GetFileName & "', Line: " & .GetFileLineNumber & vbCrLf
            End With
          Next i

        Else

          'lstrCallStack += lpException.GetType.ToString & ": " & lpException.Message & " at '" & lpException.Source & vbCrLf
          'lstrCallStack += lpException.GetType.ToString & ": " & lpException.Message & lpException.StackTrace
          lobjCallStackBuilder.AppendFormat("{0}: {1}{2}", lpException.GetType.ToString, lpException.Message,
                                              lpException.StackTrace)
          Dim lobjInnerException As Exception
          Dim lobjPreviousInnerException As Exception
          lobjInnerException = lpException.InnerException
          lobjPreviousInnerException = lobjInnerException

          Do Until lobjInnerException Is Nothing
            'lstrCallStack += lobjInnerException.Message & " at '" & lobjInnerException.Source & vbCrLf
            lobjCallStackBuilder.AppendFormat("{0} at '{1}'{2}", lobjInnerException.Message, lobjInnerException.Source,
                                                ControlChars.CrLf)
            lobjPreviousInnerException = lobjInnerException
            lobjInnerException = lobjInnerException.InnerException
          Loop

          If lobjPreviousInnerException IsNot Nothing Then
            'lstrCallStack += lobjPreviousInnerException.StackTrace
            lobjCallStackBuilder.Append(lobjPreviousInnerException.StackTrace)
          End If

        End If

        'Return lstrCallStack
        Return lobjCallStackBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Helper::FormatCallStack")
        Return String.Empty
      End Try
    End Function

    ''' <summary>
    '''     Recurses the inner exception stack of the parent exception and looks 
    ''' for an inner exception matching the type requested.
    ''' </summary>
    ''' <param name="lpParentException" type="System.Exception">
    '''     <para>
    '''         The parent exception to check.
    '''     </para>
    ''' </param>
    ''' <param name="lpInnerExceptionType" type="System.Type">
    '''     <para>
    '''         The type of exception to look for.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     The first occurance of the target exception in the inner exception stack, 
    ''' starting from the top down.  If there is no inner exception or there are no 
    ''' child exceptions matching the requested type, this method returns null.
    ''' </returns>
    Public Shared Function GetInnerException(lpParentException As Exception, lpInnerExceptionType As Type) As Exception
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpParentException)
        ArgumentNullException.ThrowIfNull(lpInnerExceptionType)
#Else
        If lpParentException Is Nothing Then
          Throw New ArgumentNullException(nameof(lpParentException))
        End If
        If lpInnerExceptionType Is Nothing Then
          Throw New ArgumentNullException(nameof(lpInnerExceptionType))
        End If
#End If

        If lpParentException.InnerException Is Nothing Then
          Return Nothing
        End If

        If lpParentException.InnerException.GetType.Name = lpInnerExceptionType.Name Then
          Return lpParentException.InnerException
        Else
          Return GetInnerException(lpParentException.InnerException, lpInnerExceptionType)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Writes detailed information on the exception to the debug window.
    ''' </summary>
    ''' <param name="Ex">The exception to write.</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteExceptionInfo(ByVal Ex As Exception)

      Try
        Debug.WriteLine("--------- Exception Data ---------")
        Debug.WriteLine(String.Format("Message: {0}", Ex.Message))
        Debug.WriteLine(String.Format("Exception Type: {0}", Ex.GetType.FullName))
        Debug.WriteLine(String.Format("Source: {0}", Ex.Source))
        Debug.WriteLine(String.Format("StackTrace: {0}", Ex.StackTrace))
        Debug.WriteLine(String.Format("TargetSite: {0}", Ex.TargetSite))
      Catch newEx As Exception
        ApplicationLogging.LogException(Ex, "Helper::WriteExceptionInfo")
      End Try
    End Sub

#End Region

    Public Shared Function GetTypeFromAssembly(lpAssembly As Assembly, lpTargetType As String) As Type
      Try
        Return GetTypeFromAssembly(lpAssembly, lpTargetType, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetTypeFromAssembly(lpAssembly As Assembly, lpTargetType As String,
                                             lpRecurseReferencedAssemblies As Boolean) As Type
      Try
        'Dim lobjTypes As Object = Nothing
        'Dim lobjInstanceTypes As New List(Of Type)
        'Dim lobjInstance As Object = Nothing

        Dim lobjReturnType As Type =
              lpAssembly.GetTypes.Where(Function(t) String.Compare(t.Name, lpTargetType, True) = 0).FirstOrDefault

        If lobjReturnType IsNot Nothing Then
          Return lobjReturnType
        Else
          If lpRecurseReferencedAssemblies = True Then
            Dim lobjReferencedAssemblies As IEnumerable(Of AssemblyName) =
                  From rAn In lpAssembly.GetReferencedAssemblies Where rAn.FullName.StartsWith("Documents") Select rAn
            For Each lobjReferencedAssemblyName As AssemblyName In lobjReferencedAssemblies
              lobjReturnType = GetTypeFromAssembly(Assembly.Load(lobjReferencedAssemblyName), lpTargetType,
                                                   lpRecurseReferencedAssemblies)
              If lobjReturnType IsNot Nothing Then
                Return lobjReturnType
              End If
            Next
          Else
            Return Nothing
          End If
        End If

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetTypesFromAssembly(lpAssembly As Assembly, lpTargetType As Type) As List(Of Type)
      Try

        Dim lobjTypes As Object = Nothing
        Dim lobjInstanceTypes As New List(Of Type)
        Dim lobjInstance As Object = Nothing

        lobjTypes = lpAssembly.GetTypes.Where(Function(t) lpTargetType.IsAssignableFrom(t))

        For Each lobjType As Type In lobjTypes
          If lobjType.IsAbstract Then
            Continue For
          End If

          If lobjType.IsInterface Then
            Continue For
          End If

          ' TODO: Figure out how to load the dependent types from other assemblies
          lobjInstance = Activator.CreateInstance(lobjType)

          If lobjInstance IsNot Nothing Then
            lobjInstanceTypes.Add(lobjType)
          End If

          lobjInstance = Nothing

        Next

        Return lobjInstanceTypes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks the specified stream to see if it represents an xml document.
    ''' </summary>
    ''' <param name="lpStream">The stream to check.</param>
    ''' <returns>True if the first line starts with 
    ''' an xml header, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsXmlStream(ByVal lpStream As Stream) As Boolean
      Try

        If (lpStream.CanSeek = False) Then
          Throw New Exception("IsXmlStream requires a stream that is seekable.")
        End If

        'Using lobjStreamReader As New StreamReader(lpStream)
        Dim lobjStreamReader As New StreamReader(lpStream)
        lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
        Dim lobjBuffer As Char()
        ReDim lobjBuffer(4)
        lobjStreamReader.ReadBlock(lobjBuffer, 0, 5)
        Dim lstrFirstFive As New String(lobjBuffer)
        lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
        If String.Equals("<?xml", lstrFirstFive, StringComparison.InvariantCultureIgnoreCase) Then
          lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
          Return True
        Else
          lobjStreamReader.BaseStream.Seek(0, SeekOrigin.Begin)
          Return False
        End If
        'End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetTypeFromName(lpTargetType As String) As Type
      Try
        For Each lobjAssembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
          'Dim lobjType As Type = lobjAssembly.GetType(lpTargetType)
          Dim lobjType As Type = GetTypeFromAssembly(lobjAssembly, lpTargetType, True)
          If lobjType IsNot Nothing Then
            Return lobjType
          Else
            Continue For
          End If
        Next
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Takes a fully qualified file path and removes 
    ''' any double back slashes (except for the front of the path).
    ''' </summary>
    ''' <param name="lpOriginalPath">The path to clean.</param>
    ''' <returns></returns>
    ''' <remarks>To handle UNC paths, double backslashes 
    ''' at the front of the string are preserved.</remarks>
    Public Shared Function RemoveExtraBackSlashesFromFilePath(ByVal lpOriginalPath As String) As String
      Try

        'If lpOriginalPath.StartsWith("\\") Then
        '  Return String.Format("\{0}", lpOriginalPath.Replace("\\", "\"))
        'Else
        '  Return lpOriginalPath.Replace("\\", "\")
        'End If

        Return CleanPath(lpOriginalPath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Helper::CleanFilePath('{0}')", lpOriginalPath))
        Return lpOriginalPath
      End Try
    End Function

#Region "File Methods"

    ''' <summary>
    ''' Ensures a valid file name by replacing invalid characters (\ / : * ? " &lt; &gt; |) with the value supplied for lpCleanChar
    ''' </summary>
    ''' <param name="lpFileName">The file name to clean</param>
    ''' <param name="lpCleanChar">The character to replace with</param>
    ''' <returns>Cleaned up file name</returns>
    ''' <remarks></remarks>
    Public Shared Function CleanFile(ByVal lpFileName As String, ByVal lpCleanChar As Char,
                                       Optional lpIsFullyQualifiedPath As Boolean = True) As String

      Try
        ' Eliminate all instances of \ / : * ? " < > |
        'Dim lstrCleanFileName As String = Path.GetFileName(lpFileName)
        'Dim lblnParameterWasFullPath As Boolean
        'Dim invalidFileNameCharacters As Char() = Path.GetInvalidFileNameChars()

        Dim lstrCleanFileName As String
        Dim lblnParameterWasFullPath As Boolean
        Dim invalidFileNameCharacters As Char() = Path.GetInvalidFileNameChars()

        ' First check for double quotes in the file name
        lpFileName = lpFileName.Replace("""", lpCleanChar)

        '' Clear out any embedded tabs before checking to see if the path is rooted
        'lpFileName = lpFileName.Replace(vbTab, lpCleanChar)

        '' Clear out any embedded pipes before checking to see if the path is rooted
        'lpFileName = lpFileName.Replace("|", lpCleanChar)

        ' Check to see if we were pased a full path
        If lpIsFullyQualifiedPath AndAlso Path.IsPathRooted(lpFileName) Then
          lblnParameterWasFullPath = True
          ' Get the actual file name
          Dim lstrFolderName As String = Path.GetDirectoryName(lpFileName)
          'lstrCleanFileName = Path.GetFileName(lpFileName)
          lstrCleanFileName = lpFileName.Replace(String.Format("{0}\", lstrFolderName), String.Empty)
        Else
          ' Make sure there are no slashes in the file name first, they will be incorrectly interpreted 
          lblnParameterWasFullPath = False
          ' Eliminate any possible embedded slashes
          lstrCleanFileName = lpFileName.Replace("\", "_").Replace("/", "_")
        End If

        For badCharCounter As Integer = 0 To invalidFileNameCharacters.Length - 1
          lstrCleanFileName = lstrCleanFileName.Replace(invalidFileNameCharacters(badCharCounter), lpCleanChar)
        Next

        If lblnParameterWasFullPath Then
          lstrCleanFileName = String.Format("{0}\{1}", Path.GetDirectoryName(lpFileName), lstrCleanFileName)
        End If

        Return lstrCleanFileName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Helper::CleanFile('{0}', '{1}')", lpFileName, lpCleanChar))
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Looks up the mime type for the supplied extension in the registry.
    ''' </summary>
    ''' <param name="lpExtension">The extension to search for.</param>
    ''' <param name="lpDefaultValue">Default value to be returned if the mime type is not found in the registry.  The standard default is text/plain.</param>
    ''' <returns>A standard mime type name from HKEY_CLASSES_ROOT\.ext where ext is the value passed in as lpExtension.</returns>
    ''' <remarks></remarks>
    Shared Function GetMIMEType(ByVal lpExtension As String) As String


      Try
        Return MimeMapping.MimeUtility.GetMimeMapping(lpExtension)
      Catch ex As Exception
        ApplicationLogging.LogException(ex,
                                        String.Format("Helper::GetMIMEType('{0}')", lpExtension))
        Return Nothing
      End Try
    End Function

#End Region

    ''' <summary>
    ''' Splits out information from the specified 
    ''' input string for the specified atribute
    ''' </summary>
    ''' <param name="lpInputString">The string to search</param>
    ''' <param name="lpAttribute">The attribute to search for</param>
    ''' <param name="lpDelimiter">The delimiter between attribute/value pairs</param>
    ''' <returns>The value of the specified attribute, if found</returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exception">
    ''' If the specified attribute id not 
    ''' found then an exception is thrown
    ''' </exception>
    Public Shared Function GetInfoFromString(ByVal lpInputString As String,
                                               ByVal lpAttribute As String,
                                               Optional ByVal lpDelimiter As String = ";") As String

      Try
        Dim lstrNameValuePairs() As String = lpInputString.Split(lpDelimiter)
        Dim lstrNameValuePair() As String
        Dim lblnInfoFound As Boolean = False

        For Each lstrPair As String In lstrNameValuePairs
          'debug.writeline(lstrPair)
          lstrNameValuePair = lstrPair.Trim.Split("=")
          If lstrNameValuePair(0) = lpAttribute Then
            lblnInfoFound = True
            Return lstrNameValuePair(1)
          End If
        Next

        If lblnInfoFound = False Then
          Throw New Exception("'" & lpAttribute & "' not found in '" & lpInputString)
        End If

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex,
                                          String.Format("Helper::GetInfoFromString('{0}', '{1}', '{2}')", lpInputString,
                                                        lpAttribute, lpDelimiter))
        Return Nothing
      End Try
    End Function

    Public Shared Function ReadTextList(lpListFilePath As String) As IList(Of String)

      Try

        Dim lobjReturnList As New List(Of String)

        ' Make sure the file exists
        VerifyFilePath(lpListFilePath, True)

        ' Create an instance of StreamReader to read from a file.
        ' The using statement also closes the StreamReader.
        Using lobjStreamReader As New StreamReader(lpListFilePath)

          Dim lstrLine As String

          ' Read and display lines from the file until the end of
          ' the file is reached.
          Do
            lstrLine = lobjStreamReader.ReadLine()

            If Not String.IsNullOrEmpty(lstrLine) Then
              lobjReturnList.Add(lstrLine)
            End If

          Loop Until lstrLine Is Nothing

        End Using

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ReadExcelList(lpListFilePath As String) As IList(Of String)

      Try

        Dim lobjReturnList As New List(Of String)

        ' Make sure the file exists
        VerifyFilePath(lpListFilePath, True)

        Dim lobjExcelReader As ExcelDataReader.IExcelDataReader = Nothing

        Using lobjFileStream As FileStream = File.Open(lpListFilePath, FileMode.Open, FileAccess.Read)

          Select Case Path.GetExtension(lpListFilePath).ToLower

            Case ".xls"
              lobjExcelReader = ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(lobjFileStream)

            Case ".xlsx"
              lobjExcelReader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(lobjFileStream)
          End Select

          If lobjExcelReader IsNot Nothing Then

            'lobjExcelReader.IsFirstRowAsColumnNames=True
            'Dim lobjDataSet As DataSet = lobjExcelReader.AsDataSet
            While lobjExcelReader.Read
              lobjReturnList.Add(lobjExcelReader.GetString(0))
            End While

            lobjExcelReader.Close()
          End If

        End Using

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ReadShort(lpStream As Stream, lpIsIntel As Boolean) As UShort
      Try
        Dim lobjByte As Byte() = New Byte(1) {}
        lpStream.Read(lobjByte, 0, lobjByte.Length)
        Return ToShort(lpIsIntel, lobjByte(0), lobjByte(1))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ToShort(lpIsIntel As Boolean, b0 As Byte, b1 As Byte) As UShort
      Try
        If lpIsIntel Then
          Return CUShort((CInt(b1) << 8) Or CInt(b0))
        Else
          Return CUShort((CInt(b0) << 8) Or CInt(b1))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

  End Class

End Namespace