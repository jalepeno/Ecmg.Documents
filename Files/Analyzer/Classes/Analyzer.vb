'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  Analyzer.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 3:52:27 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports Documents.Utilities
Imports Microsoft.VisualBasic.CompilerServices

#End Region

Namespace Files

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Analyzer
    Implements IDisposable

#Region "Class Constants"

    Private Const TEN_K As Integer = 10240
    Private Const HALF_MB As Integer = 524288
    Private Const HALF_MB_PLUS_ONE_BYTE As Integer = 524289
    Private Const ONE_MB As Integer = 1048576
    Private Const ONE_MB_PLUS_TWO_BYTES As Integer = 1048578
    Private Const FOUR_K As Integer = 4096
    Private Const TEN_MB As Integer = 10485760
    Private Const FIVE_MB As Integer = 5242880
    Private Const TEN_MB_PLUS_TWO_BYTES As Integer = 10485762
    Private Const FIVE_MB_PLUS_ONE_BYTE As Integer = 5242881

#End Region

#Region "Class Variables"

    Private mobjFileDefinitions As New FileDefinitions
    Private mobjFoundDefinitions As New FileDefinitions
    Private mintFileFrontSize As Integer
    Private mintFileBackSize As Integer
    Private mobjResults As IFileTestResults = New FileTestResults
    Private mobjFrontBlock As ByteString
    Private mstrSubmittedFile As String
    Private menuPrecision As Precision
    Private mintSampleSize As Integer
    Private mblnIgnoreZeroByteFiles As Boolean = False

    Private mobjADefinitions As IFileDefinitions
    Private mobjBDefinitions As IFileDefinitions
    Private mobjCDefinitions As IFileDefinitions

    'Private mobjLogSession As Gurock.SmartInspect.Session = Nothing

    Private Shared mobjInstance As Analyzer
    Private Shared mintReferenceCount As Integer

#End Region

#Region "Constructors"

    Private Sub New()
      Try
        'InitializeLogSession()
        InitializeDefinitions()
        Precision = Precision.Quick
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    Public Property IgnoreZeroByteFiles As Boolean
      Get
        Try
          Return mblnIgnoreZeroByteFiles
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          mblnIgnoreZeroByteFiles = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

    Public Property FileDefinitions As FileDefinitions
      Get
        Try
          Return mobjFileDefinitions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As FileDefinitions)
        Try
          mobjFileDefinitions = value
          If mintFileFrontSize < mobjFileDefinitions.FrontBlockSize Then
            mintFileFrontSize = mobjFileDefinitions.FrontBlockSize
          End If
          mobjADefinitions = mobjFileDefinitions.ItemsByPopularity(Rating.A)
          mobjBDefinitions = mobjFileDefinitions.ItemsByPopularity(Rating.B)
          mobjCDefinitions = mobjFileDefinitions.ItemsByPopularity(Rating.C)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property FoundDefinitions As FileDefinitions
      Get
        Try
          Return mobjFoundDefinitions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session
    '  Get
    '    Try
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    Public Property Precision As Precision
      Get
        Try
          Return menuPrecision
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Precision)
        Try
          menuPrecision = value
          ' Set the sample size based on the requested precision
          Select Case Precision
            Case Precision.Accurate
              mintSampleSize = TEN_MB
            Case Precision.Quick
              mintSampleSize = ONE_MB
            Case Precision.Rocket
              mintSampleSize = TEN_K
          End Select
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property Results As IFileTestResults
      Get
        Try
          Return mobjResults
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    'Public Sub Analyze()
    '  Try

    '    Dim lintNumberThree As Integer
    '    Dim lobjResult As FileTestResult
    '    Dim lblnFlag As Boolean = True
    '    Dim lintNumberEight As Integer = 0
    '    mobjResults.Clear()

    '    Dim lobjEnumerator As IEnumerator = Me.mobjFileDefinitions.GetEnumerator

    '    Do While lobjEnumerator.MoveNext
    '      Dim lobjCurrentDefinition As IFileDefinition = DirectCast(lobjEnumerator.Current, IFileDefinition)
    '      Dim lintNumberFive As Integer = 0

    '      If (Me.mintFileFrontSize >= lobjCurrentDefinition.FrontBlockSize) Then
    '        Dim lobjEnumeratorTwo As IEnumerator = lobjCurrentDefinition.Patterns.GetEnumerator

    '        Do While lobjEnumeratorTwo.MoveNext
    '          Dim lobjTestPattern As IHeaderPattern = DirectCast(lobjEnumeratorTwo.Current, IHeaderPattern)
    '          Dim lintNumberFour As Integer = 0
    '          Dim lintNumberEleven As Integer = (lobjTestPattern.Length - 1)
    '          lintNumberThree = 0

    '          Do While (lintNumberThree <= lintNumberEleven)

    '            If (Me.mobjFrontBlock.data((lobjTestPattern.Position + lintNumberThree)) = lobjTestPattern.Pattern(lintNumberThree)) Then
    '              lintNumberFour += 1
    '            End If

    '            lintNumberThree += 1
    '          Loop

    '          If (lintNumberFour = lobjTestPattern.Length) Then
    '            lintNumberFive = (lintNumberFive + (lobjTestPattern.Length * lobjTestPattern.Points))

    '          Else
    '            lintNumberFive = 0
    '            Exit Do
    '          End If

    '        Loop

    '        lobjEnumeratorTwo.Reset()

    '        If (lintNumberFive > 0) Then

    '          If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then

    '            Dim lobjBuffer As Byte() = Nothing

    '            If lblnFlag Then

    '              Dim lobjReader As BinaryReader = Nothing
    '              Dim lobjStream As FileStream = Nothing

    '              Try
    '                lobjStream = New FileStream(Me.mstrSubmittedFile, FileMode.Open, FileAccess.Read)
    '                lobjReader = New BinaryReader(lobjStream)

    '                If (lobjStream.Length <> 0) Then

    '                  If (lobjStream.Length <= &HA00000) Then
    '                    lobjBuffer = lobjReader.ReadBytes(CInt(lobjStream.Length))

    '                  Else
    '                    lobjBuffer = lobjReader.ReadBytes(&H500000)
    '                    lobjBuffer = DirectCast(Utils.CopyArray(DirectCast(lobjBuffer, Array), New Byte(&HA00002 - 1) {}), Byte())
    '                    lobjStream.Seek(-5242880, SeekOrigin.End)

    '                    Dim lobjBufferTwo As Byte() = lobjReader.ReadBytes(&H500000)
    '                    lobjBuffer(&H500000) = &H7C
    '                    lobjBufferTwo.CopyTo(lobjBuffer, &H500001)
    '                  End If

    '                  lblnFlag = False
    '                  FileDefinition.ByteAlphaToUpper(lobjBuffer)
    '                End If

    '              Catch exception1 As Exception
    '                ProjectData.SetProjectError(exception1)
    '                Me.mobjResults.Clear()
    '                Throw New Exception("ERROR")

    '              Finally

    '                If (Not lobjStream Is Nothing) Then
    '                  lobjStream.Close()

    '                  If (Not lobjReader Is Nothing) Then
    '                    lobjReader.Close()
    '                  End If

    '                End If

    '              End Try

    '            End If

    '            lobjEnumeratorTwo = lobjCurrentDefinition.GlobalStrings.GetEnumerator

    '            Do While lobjEnumeratorTwo.MoveNext

    '              Dim str As ByteString = DirectCast(lobjEnumeratorTwo.Current, ByteString)

    '              If (Me.ByteArraySearchCS(lobjBuffer, str.data) <> -1) Then
    '                lintNumberFive = (lintNumberFive + (str.data.Length * 500))

    '              Else
    '                lintNumberFive = 0
    '                Exit Do
    '              End If

    '            Loop

    '            lobjEnumeratorTwo.Reset()
    '          End If

    '          If (lintNumberFive > 0) Then
    '            lobjResult = New FileTestResult
    '            lobjResult.FileType = lobjCurrentDefinition.FileInfo.FileType
    '            lobjResult.FileExt = lobjCurrentDefinition.FileInfo.Extension
    '            'lobjResult.ExtraInfo = lobjCurrentDefinition.ExtraInfo
    '            lobjResult.ExtraInfo.FilePts = (StringType.FromInteger(lintNumberFive) & "/" & StringType.FromInteger(lobjCurrentDefinition.Patterns.Count))

    '            If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then
    '              lobjResult.ExtraInfo.FilePts = (lobjResult.ExtraInfo.FilePts & "/" & StringType.FromInteger(lobjCurrentDefinition.GlobalStrings.Count))
    '            End If

    '            lobjResult.FileType = String.Concat(New String() {lobjResult.FileType, " (", StringType.FromInteger(lintNumberFive), "/", StringType.FromInteger(lobjCurrentDefinition.Patterns.Count), ")"})
    '            lobjResult.Points = lintNumberFive
    '            Me.mobjResults.Add(lobjResult)
    '            lintNumberEight = (lintNumberEight + lintNumberFive)
    '          End If

    '        End If
    '      End If

    '    Loop

    '    lobjEnumerator.Reset()

    '    If (Me.mobjResults.Count > 0) Then

    '      Dim lintNumberTen As Integer = (Me.mobjResults.Count - 1)
    '      lintNumberThree = 0

    '      Do While (lintNumberThree <= lintNumberTen)
    '        lobjResult = DirectCast(Me.mobjResults.Item(lintNumberThree), FileTestResult)
    '        lobjResult.Perc = CSng((CDbl((lobjResult.Points * 100)) / CDbl(lintNumberEight)))
    '        Me.mobjResults.Item(lintNumberThree) = lobjResult
    '        lintNumberThree += 1
    '      Loop

    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Public Function Analyze() As ITestResults
    '  Try

    '    Dim lintCounter As Integer
    '    Dim lobjResults As New TestResults
    '    Dim lobjResult As TestResult
    '    Dim lblnFlag As Boolean = True
    '    Dim lintPercentage As Integer = 0
    '    mobjResults.Clear()

    '    Dim lintHexSize As Integer

    '    For Each lobjCurrentDefinition As IFileDefinition In mobjFileDefinitions

    '      Dim lintPoints As Integer = 0

    '      If (Me.mintFileFrontSize >= lobjCurrentDefinition.FrontBlockSize) Then

    '        For Each lobjTestPattern As IHeaderPattern In lobjCurrentDefinition.Patterns

    '          Dim lintMatchCount As Integer = 0
    '          Dim lintPatternLength As Integer = (lobjTestPattern.Length - 1)
    '          lintCounter = 0

    '          Do While (lintCounter <= lintPatternLength)

    '            If (Me.mobjFrontBlock.data((lobjTestPattern.Position + lintCounter)) = lobjTestPattern.Pattern(lintCounter)) Then
    '              lintMatchCount += 1
    '            End If

    '            lintCounter += 1
    '          Loop

    '          If (lintMatchCount = lobjTestPattern.Length) Then
    '            lintPoints = (lintPoints + (lobjTestPattern.Length * lobjTestPattern.Points))

    '          Else
    '            lintPoints = 0
    '            Exit For
    '          End If

    '        Next

    '        If (lintPoints > 0) Then

    '          If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then

    '            Dim lobjBuffer As Byte() = Nothing

    '            If lblnFlag Then

    '              Dim lobjReader As BinaryReader = Nothing
    '              Dim lobjStream As FileStream = Nothing

    '              Try
    '                lobjStream = New FileStream(Me.mstrSubmittedFile, FileMode.Open, FileAccess.Read)
    '                lobjReader = New BinaryReader(lobjStream)

    '                If (lobjStream.Length <> 0) Then

    '                  If (lobjStream.Length <= TEN_MB) Then
    '                    lobjBuffer = lobjReader.ReadBytes(CInt(lobjStream.Length))

    '                  Else
    '                    lobjBuffer = lobjReader.ReadBytes(FIVE_MB)
    '                    lobjBuffer = DirectCast(Utils.CopyArray(DirectCast(lobjBuffer, Array), New Byte(TEN_MB_PLUS_TWO_BYTES - 1) {}), Byte())
    '                    lobjStream.Seek(-FIVE_MB, SeekOrigin.End)

    '                    Dim lobjBufferTwo As Byte() = lobjReader.ReadBytes(FIVE_MB)
    '                    lobjBuffer(FIVE_MB) = 124
    '                    lobjBufferTwo.CopyTo(lobjBuffer, FIVE_MB_PLUS_ONE_BYTE)
    '                  End If

    '                  lblnFlag = False
    '                  FileDefinition.ByteAlphaToUpper(lobjBuffer)
    '                End If

    '              Catch exception1 As Exception
    '                ProjectData.SetProjectError(exception1)
    '                Me.mobjResults.Clear()
    '                Throw New Exception("ERROR")

    '              Finally

    '                If (Not lobjStream Is Nothing) Then
    '                  lobjStream.Close()

    '                  If (Not lobjReader Is Nothing) Then
    '                    lobjReader.Close()
    '                  End If

    '                End If

    '              End Try

    '            End If

    '            For Each lstrGlobalString As ByteString In lobjCurrentDefinition.GlobalStrings

    '              If (Me.ByteArraySearchCS(lobjBuffer, lstrGlobalString.data) <> -1) Then
    '                lintPoints = (lintPoints + (lstrGlobalString.data.Length * 500))

    '              Else
    '                lintPoints = 0
    '                Exit For
    '              End If

    '            Next

    '          End If

    '          If (lintPoints > 0) Then
    '            lobjResult = New TestResult
    '            lobjResult.FileType = lobjCurrentDefinition.FileInfo.FileType
    '            lobjResult.Extension = lobjCurrentDefinition.FileInfo.Extension
    '            lobjResult.FilePoints = String.Format("{0}/{1}", lintPoints, lobjCurrentDefinition.Patterns.Count)

    '            If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then
    '              lobjResult.FilePoints = String.Format("{0}/{1}", lobjResult.Points, lobjCurrentDefinition.GlobalStrings.Count)
    '            End If

    '            lobjResult.FileType = String.Format("{0} ({1}/{2})", lobjResult.FileType, lintPoints, lobjCurrentDefinition.Patterns.Count)
    '            lobjResult.Points = lintPoints
    '            lobjResults.Add(lobjResult)
    '            lintPercentage = (lintPercentage + lintPoints)
    '          End If

    '        End If
    '      End If

    '    Next

    '    If (lobjResults.Count > 0) Then
    '      For Each lobjTestResult As ITestResult In lobjResults
    '        lobjTestResult.Percentage = ((lobjTestResult.Points * 100) / lintPercentage)
    '      Next
    '    End If

    '    mobjResults = lobjResults

    '    Return lobjResults

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function Analyze(lpFileName As String) As ITestResults
    '  Try
    '    If File.Exists(lpFileName) Then

    '      Dim lobjReader As BinaryReader = Nothing
    '      Dim lobjStream As FileStream = Nothing
    '      Me.mintFileFrontSize = FOUR_K

    '      Dim lobjBuffer As Byte()

    '      Try
    '        lobjStream = New FileStream(lpFileName, FileMode.Open, FileAccess.Read)
    '        lobjReader = New BinaryReader(lobjStream)

    '        If (lobjStream.Length <> 0) Then

    '          If (lobjStream.Length < Me.mintFileFrontSize) Then
    '            Me.mintFileFrontSize = CInt(lobjStream.Length)
    '          End If

    '          Me.mobjFrontBlock.data = lobjReader.ReadBytes(Me.mintFileFrontSize)

    '          If lobjStream.CanSeek Then
    '            lobjStream.Position = 0
    '          End If

    '          If (lobjStream.Length <= TEN_MB) Then
    '            lobjBuffer = lobjReader.ReadBytes(CInt(lobjStream.Length))

    '          Else
    '            lobjBuffer = lobjReader.ReadBytes(FIVE_MB)
    '            lobjBuffer = DirectCast(Utils.CopyArray(DirectCast(lobjBuffer, Array), New Byte(TEN_MB_PLUS_TWO_BYTES - 1) {}), Byte())
    '            lobjStream.Seek(-FIVE_MB, SeekOrigin.End)

    '            Dim lobjBufferTwo As Byte() = lobjReader.ReadBytes(FIVE_MB)
    '            lobjBuffer(FIVE_MB) = 124
    '            lobjBufferTwo.CopyTo(lobjBuffer, FIVE_MB_PLUS_ONE_BYTE)
    '          End If

    '          Me.mstrSubmittedFile = lpFileName
    '          Return AnalyzeBytes(lobjBuffer)

    '        Else
    '          Throw New InvalidOperationException("Cannot analyze a zero byte file.")
    '        End If

    '      Catch exception1 As Exception
    '        ProjectData.SetProjectError(exception1)
    '        Me.mintFileFrontSize = 0
    '        'Throw New Exception("ERROR")
    '        ' Re-throw the exception to the caller
    '        Throw
    '      Finally

    '        If (Not lobjStream Is Nothing) Then
    '          lobjStream.Close()

    '          If (Not lobjReader Is Nothing) Then
    '            lobjReader.Close()
    '          End If

    '        End If

    '      End Try

    '    Else
    '      Throw New InvalidOperationException("Invalid file path.")
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#Region "Singleton Support"

    Public Shared ReadOnly Property Instance As Analyzer
      Get
        Try
          If Not Helper.CallStackContainsMethodName("AssignObjectProperties") Then
            Return GetInstance(False)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property


    Friend Shared Function GetInstance(ByVal lpForceRefresh As Boolean) As Analyzer
      Try
        If lpForceRefresh = True Then
          mobjInstance = New Analyzer()
        ElseIf mobjInstance Is Nothing OrElse mobjInstance.IsDisposed = True Then
          mobjInstance = New Analyzer()
        End If
        mintReferenceCount += 1
        Return mobjInstance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Public Function Analyze(lpFileName As String) As IFileTestResults
      Try
        If File.Exists(lpFileName) Then

          Dim lobjReader As BinaryReader = Nothing
          Dim lobjStream As FileStream = Nothing
          Me.mintFileFrontSize = FOUR_K
          Dim lobjBuffer As Byte()

          Try
            lobjStream = New FileStream(lpFileName, FileMode.Open, FileAccess.Read)
            lobjReader = New BinaryReader(lobjStream)

            If (lobjStream.Length <> 0) Then

              ' Set the sample size based on the requested precision
              If Precision = Precision.Complete Then
                mintSampleSize = lobjStream.Length
              End If

              If (lobjStream.Length < Me.mintFileFrontSize) Then
                Me.mintFileFrontSize = CInt(lobjStream.Length)
              End If

              Me.mobjFrontBlock.data = lobjReader.ReadBytes(Me.mintFileFrontSize)

              If lobjStream.CanSeek Then
                lobjStream.Position = 0
              End If

              If (lobjStream.Length <= mintSampleSize) Then
                lobjBuffer = lobjReader.ReadBytes(CInt(lobjStream.Length))

              Else
                lobjBuffer = lobjReader.ReadBytes((mintSampleSize / 2))
                lobjBuffer = DirectCast(Utils.CopyArray(DirectCast(lobjBuffer, Array), New Byte((mintSampleSize + 2) - 1) {}), Byte())
                lobjStream.Seek(-(mintSampleSize / 2), SeekOrigin.End)

                Dim lobjBufferTwo As Byte() = lobjReader.ReadBytes((mintSampleSize / 2))
                lobjBuffer((mintSampleSize / 2)) = 124
                lobjBufferTwo.CopyTo(lobjBuffer, CInt((mintSampleSize / 2)) + 1)
              End If

              Me.mstrSubmittedFile = lpFileName
              Return AnalyzeBytes(lobjBuffer)

            Else
              If IgnoreZeroByteFiles Then
                Return New FileTestResults
              Else
                Throw New InvalidOperationException("Cannot analyze a zero byte file.")
              End If
            End If

          Catch exception1 As Exception
            ProjectData.SetProjectError(exception1)
            Me.mintFileFrontSize = 0
            'Throw New Exception("ERROR")
            ' Re-throw the exception to the caller
            Throw
          Finally

            If (Not lobjStream Is Nothing) Then
              lobjStream.Close()

              If (Not lobjReader Is Nothing) Then
                lobjReader.Close()
              End If

            End If

          End Try

        Else
          Throw New InvalidOperationException("Invalid file path.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Function Analyze(lpContent As Stream) As IFileTestResults
      Try
        Me.mintFileFrontSize = FOUR_K

        Dim lobjBuffer(mintFileFrontSize - 1) As Byte
        Dim lintBytesRead As Integer

        If Not lpContent.CanRead Then
          Throw New InvalidOperationException("Cannot analyze a closed stream.")
        End If

        If lpContent.CanSeek Then
          lpContent.Position = 0
        Else
          Throw New InvalidOperationException("Stream is not seekable.")
        End If

        If (lpContent.Length <> 0) Then

          ' Set the sample size based on the requested precision
          If Precision = Precision.Complete Then
            mintSampleSize = lpContent.Length
          End If

          If (lpContent.Length < Me.mintFileFrontSize) Then
            Me.mintFileFrontSize = CInt(lpContent.Length)
            ReDim lobjBuffer(mintFileFrontSize - 1)
          End If
          lintBytesRead = lpContent.Read(lobjBuffer, 0, Me.mintFileFrontSize)
          Me.mobjFrontBlock.data = lobjBuffer

          lpContent.Position = 0

          If (lpContent.Length <= mintSampleSize) Then
            'lobjBuffer = lobjReader.ReadBytes(CInt(lobjStream.Length))
            ReDim lobjBuffer(lpContent.Length - 1)
            lintBytesRead = lpContent.Read(lobjBuffer, 0, lpContent.Length)

          Else
            'lobjBuffer = lobjReader.ReadBytes(FIVE_MB)
            ReDim lobjBuffer(mintSampleSize - 1)
            lintBytesRead = lpContent.Read(lobjBuffer, 0, (mintSampleSize / 2))
            lobjBuffer = DirectCast(Utils.CopyArray(DirectCast(lobjBuffer, Array), New Byte((mintSampleSize + 2) - 1) {}), Byte())
            lpContent.Seek(-(mintSampleSize / 2), SeekOrigin.End)

            'Dim lobjBufferTwo As Byte() = lobjReader.ReadBytes(FIVE_MB)
            Dim lobjBufferTwo((mintSampleSize / 2) - 1) As Byte
            lintBytesRead = lpContent.Read(lobjBufferTwo, 0, (mintSampleSize / 2))
            lobjBuffer((mintSampleSize / 2)) = 124
            lobjBufferTwo.CopyTo(lobjBuffer, CInt((mintSampleSize / 2) + 1))
          End If

          lpContent.Position = 0

          Return AnalyzeBytes(lobjBuffer)

        Else
          Throw New InvalidOperationException("Cannot analyze a zero byte stream.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Analyze(lpContent As Byte()) As IFileTestResults
      Try
        If (lpContent.Length <> 0) Then
          Dim lobjContentStream As New MemoryStream(lpContent)
          Return Analyze(lobjContentStream)
        Else
          Throw New InvalidOperationException("Cannot analyze empty byte array.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetMimeType(lpExtension As String, lpDefaultValue As String) As String
      Try
        ''LogSession.LogVerbose("Starting GetMimeType for extension '{0}' with a default value of '{1}'.", lpExtension, lpDefaultValue)
        'Return mobjFileDefinitions.GetMimeType(lpExtension, lpDefaultValue)
        Dim lstrTempDefault As String = "temp/plain"

        ' First look at the most popular file types
        Dim lstrMimeType As String = mobjADefinitions.GetMimeType(lpExtension, lstrTempDefault)

        If Not String.Equals(lstrMimeType, lstrTempDefault) Then
          ''LogSession.LogVerbose("GetMimeTypeA: Mimetype '{0}' found for extension '{1}'.", lstrMimeType, lpExtension)
          Return lstrMimeType
        Else
          ' If we did not find anything then test against the next most popular types.
          lstrMimeType = mobjBDefinitions.GetMimeType(lpExtension, lstrTempDefault)
          If Not String.Equals(lstrMimeType, lstrTempDefault) Then
            ''LogSession.LogVerbose("GetMimeTypeB: Mimetype '{0}' found for extension '{1}'.", lstrMimeType, lpExtension)
            Return lstrMimeType
          Else
            ' If we still did not find anything then test against the least most popular types.
            lstrMimeType = mobjCDefinitions.GetMimeType(lpExtension, lpDefaultValue)
            ''LogSession.LogVerbose("GetMimeTypeC: Mimetype '{0}' found for extension '{1}'.", lstrMimeType, lpExtension)
            Return lstrMimeType
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function IsBinary() As Boolean
      Try
        If (StringType.StrCmp(Me.mstrSubmittedFile, "", False) <> 0) Then

          Dim lintFileFrontSize As Integer = (Me.mintFileFrontSize - 1)
          Dim lintCounter As Integer = 0

          Do While (lintCounter <= lintFileFrontSize)

            Dim lobjByte As Byte = Me.mobjFrontBlock.data(lintCounter)

            If ((lobjByte < 9) Or (lobjByte > &H7E)) Then
              Return True
            End If

            lintCounter += 1
          Loop

        End If

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SubmitFile(ByVal lpFileName As String)
      Try
        If File.Exists(lpFileName) Then

          Dim lobjReader As BinaryReader = Nothing
          Dim lobjStream As FileStream = Nothing
          Me.mintFileFrontSize = &H1000

          Try
            lobjStream = New FileStream(lpFileName, FileMode.Open, FileAccess.Read)
            lobjReader = New BinaryReader(lobjStream)

            If (lobjStream.Length <> 0) Then

              If (lobjStream.Length < Me.mintFileFrontSize) Then
                Me.mintFileFrontSize = CInt(lobjStream.Length)
              End If

              Me.mobjFrontBlock.data = lobjReader.ReadBytes(Me.mintFileFrontSize)
            End If

            Me.mstrSubmittedFile = lpFileName

          Catch exception1 As Exception
            ProjectData.SetProjectError(exception1)
            Me.mintFileFrontSize = 0
            Throw New Exception("ERROR")

          Finally

            If (Not lobjStream Is Nothing) Then
              lobjStream.Close()

              If (Not lobjReader Is Nothing) Then
                lobjReader.Close()
              End If

            End If

          End Try

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        If FileDefinitions.Count > 0 Then
          lobjIdentifierBuilder.AppendFormat("{0} File Definitions", FileDefinitions.Count)
        Else
          lobjIdentifierBuilder.Append("No File Definitions")
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

#Region "Private Methods"

    'Private Sub InitializeLogSession()
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, Drawing.Color.Bisque)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Private Sub FinalizeLogSession()
    '  Try
    '    'LogSession.LogMessage("{0} disposing", Me.GetType.Name)
    '    ApplicationLogging.FinalizeLogSession(LogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Private Sub InitializeDefinitions()
      Try
        ''LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
        ''LogSession.LogMessage("Initializing Analyzer definitions")
        Dim lobjWatch As Stopwatch = Stopwatch.StartNew()

        Dim lobjStream As Stream = GetFileDefinitionStream()

        Dim lobjFileDefinitions As FileDefinitions = FileDefinitions.FromArchive(lobjStream)

        If lobjFileDefinitions IsNot Nothing Then
          FileDefinitions = lobjFileDefinitions
          lobjWatch.Stop()
          ''LogSession.LogMessage(String.Format("Initialized Analyzer definitions in {0}ms.", lobjWatch.ElapsedMilliseconds))
          ApplicationLogging.LogInformation($"Initialized Analyzer definitions in {lobjWatch.ElapsedMilliseconds}ms.")
          lobjWatch = Nothing
        Else
          Throw New InvalidOperationException("Failed to initialize file definitions from resource.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      Finally
        ''LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Sub

    Private Function GetFileDefinitionStream() As Stream
      Try
        Dim lobjResourceStream As Stream = Nothing
        Dim lobjAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim lobjResourceNames As String() = lobjAssembly.GetManifestResourceNames()

        For Each lstrResourceName As String In lobjResourceNames
          If lstrResourceName.Contains(".FileDefinitions.zip") Then
            lobjResourceStream = lobjAssembly.GetManifestResourceStream(lstrResourceName)
            Exit For
          End If
        Next

        Return lobjResourceStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function AnalyzeBytes(lpContent As Byte()) As IFileTestResults
      Try

        'If FoundDefinitions.Count = 0 Then
        '  Return AnalyzeBytes(lpContent, Nothing)
        'Else
        '  Dim lobjResults As ITestResults
        '  lobjResults = AnalyzeBytes(lpContent, FoundDefinitions)
        '  If lobjResults.Count = 0 Then
        '    lobjResults = AnalyzeBytes(lpContent, Nothing)
        '  End If
        '  Return lobjResults
        'End If

        Return AnalyzeBytes(lpContent, Nothing)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function AnalyzeBytes(lpContent As Byte(), Optional lpDefinitions As IFileDefinitions = Nothing) As IFileTestResults
      Try
        ''LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
        Dim lobjResults As New FileTestResults
        Dim lobjDefinitions As IFileDefinitions
        mobjResults.Clear()

        If lpDefinitions IsNot Nothing Then
          lobjDefinitions = lpDefinitions
        Else
          lobjDefinitions = mobjFileDefinitions
        End If

        ' Test against the most popular definitions.
        'Console.Write("A")
        lobjResults = AnalyzeAgainstDefinitions(lpContent, mobjADefinitions)

        ' If we did not find anything then test against the next most popular definitions.
        If lobjResults.Count = 0 Then
          ' Console.Write(", B")
          lobjResults = AnalyzeAgainstDefinitions(lpContent, mobjBDefinitions)
        End If

        ' If we still did not find anything then test against the least most popular definitions.
        If lobjResults.Count = 0 Then
          'Console.WriteLine(", C")
          lobjResults = AnalyzeAgainstDefinitions(lpContent, mobjCDefinitions)
        End If

        If lobjResults.Count > 0 Then
          ''LogSession.LogMessage("Primary result: {0}", lobjResults.PrimaryResult.ToString())
        End If

        Return lobjResults

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      Finally
        ''LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Function

    Private Function AnalyzeAgainstDefinitions(lpContent As Byte(), lpDefinitions As IFileDefinitions) As IFileTestResults

      Dim lobjDebugDefinition As IFileDefinition
      Dim lobjDebugTestPattern As IHeaderPattern
      Dim lintCounter As Integer
      Dim lintRequestedDataIndex As Integer

      Try

        ''LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
        Dim lobjResults As New FileTestResults
        Dim lobjResult As FileTestResult
        Dim lblnFlag As Boolean = True
        Dim lintPercentage As Integer = 0

        For Each lobjCurrentDefinition As IFileDefinition In lpDefinitions
          lobjDebugDefinition = lobjCurrentDefinition
          Dim lintPoints As Integer = 0

          If (Me.mintFileFrontSize >= lobjCurrentDefinition.FrontBlockSize) Then

            For Each lobjTestPattern As IHeaderPattern In lobjCurrentDefinition.Patterns
              lobjDebugTestPattern = lobjTestPattern

              Dim lintMatchCount As Integer = 0
              Dim lintPatternLength As Integer = (lobjTestPattern.Length - 1)
              lintCounter = 0

              Do While (lintCounter <= lintPatternLength)
                lintRequestedDataIndex = lobjTestPattern.Position + lintCounter
                If lintRequestedDataIndex > Me.mobjFrontBlock.data.Length - 1 Then
                  Exit Do
                End If
                If (Me.mobjFrontBlock.data(lintRequestedDataIndex) = lobjTestPattern.Pattern(lintCounter)) Then
                  lintMatchCount += 1
                End If

                lintCounter += 1
              Loop

              If (lintMatchCount = lobjTestPattern.Length) Then
                lintPoints = (lintPoints + (lobjTestPattern.Length * lobjTestPattern.Points))

              Else
                lintPoints = 0
                Exit For
              End If

            Next

            If (lintPoints > 0) Then

              If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then

                If lblnFlag Then

                  Try

                    lblnFlag = False
                    ByteStringFactory.ByteAlphaToUpper(lpContent)
                    'End If

                  Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Me.mobjResults.Clear()
                    Throw New Exception("ERROR")

                  End Try

                End If

                For Each lstrGlobalString As ByteString In lobjCurrentDefinition.GlobalStrings

                  If (Me.ByteArraySearchCS(lpContent, lstrGlobalString.data) <> -1) Then
                    lintPoints = (lintPoints + (lstrGlobalString.data.Length * 500))

                  Else
                    lintPoints = 0
                    Exit For
                  End If

                Next

              End If

              If (lintPoints > 0) Then
                lobjResult = New FileTestResult(lobjCurrentDefinition)
                lobjResult.FilePoints = String.Format("{0}/{1}", lintPoints, lobjCurrentDefinition.Patterns.Count)

                If (lobjCurrentDefinition.GlobalStrings.Count > 0) Then
                  lobjResult.FilePoints = String.Format("{0}/{1}", lobjResult.Points, lobjCurrentDefinition.GlobalStrings.Count)
                End If

                lobjResult.FileType = String.Format("{0} ({1}/{2})", lobjResult.FileType, lintPoints, lobjCurrentDefinition.Patterns.Count)
                lobjResult.Points = lintPoints
                lobjResults.Add(lobjResult)
                lobjCurrentDefinition.ItemsFound += 1
                If lpDefinitions Is Nothing Then
                  FoundDefinitions.Add(lobjCurrentDefinition)
                End If
                lintPercentage = (lintPercentage + lintPoints)
              End If

            End If
          End If

        Next

        If (lobjResults.Count > 0) Then
          For Each lobjTestResult As IFileTestResult In lobjResults
            lobjTestResult.Percentage = ((lobjTestResult.Points * 100) / lintPercentage)
          Next
          'lobjResults.Sort()
        End If

        lobjResults.OrderByPercentage()

        mobjResults = lobjResults

        Return lobjResults

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      Finally
        ''LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Function

    Private Function ByteArraySearchCS(ByVal lpBig As Byte(),
                                     ByVal lpSearch As Byte()) As Integer

      Try
        Dim lintNumberArray As Integer() = New Integer(&H100 - 1) {}
        Dim lintSearchSize As Integer = lpSearch.Length
        Dim lintBigSize As Integer = lpBig.Length
        Dim lintIndex As Integer = 0

        Do
          lintNumberArray(lintIndex) = lintSearchSize
          lintIndex += 1
        Loop While (lintIndex <= &HFF)

        Dim lintNumberSeven As Integer = lintSearchSize
        lintIndex = 1

        Do While (lintIndex <= lintNumberSeven)

          Dim lobjByte As Byte = lpSearch((lintIndex - 1))
          lintNumberArray(lobjByte) = (lintSearchSize - lintIndex)
          lintIndex += 1
        Loop

        Dim lintNumberThree As Integer = lintSearchSize
        lintIndex = lintSearchSize

        Do

          If (lpBig((lintNumberThree - 1)) = lpSearch((lintIndex - 1))) Then
            lintNumberThree -= 1
            lintIndex -= 1

          Else

            If (((lintSearchSize - lintIndex) + 1) > lintNumberArray(lpBig((lintNumberThree - 1)))) Then
              lintNumberThree = (((lintNumberThree + lintSearchSize) - lintIndex) + 1)

            Else
              lintNumberThree = (lintNumberThree + lintNumberArray(lpBig((lintNumberThree - 1))))
            End If

            lintIndex = lintSearchSize
          End If

        Loop While ((lintIndex >= 1) And (lintNumberThree <= lintBigSize))

        If (lintNumberThree >= lintBigSize) Then
          Return -1
        End If

        Return lintNumberThree
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          FileDefinitions.Clear()
          FoundDefinitions.Clear()
          Results.Clear()
          'FinalizeLogSession()
        End If

        ' DISPOSETODO: free your own state (unmanaged objects).
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

  End Class

End Namespace