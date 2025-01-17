'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions

#End Region

Namespace Utilities

  ''' <summary>
  ''' Container object for the three types of versions tracked for an assembly
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class AssemblyVersionInfo
    Implements IEqualityComparer
    Implements IEquatable(Of AssemblyVersionInfo)

#Region "Class Constants"

    Private Const ASSEMBLY_VERSION As String = "AssemblyVersion"
    Private Const ASSEMBLY_INFORMATIONAL_VERSION As String = "AssemblyInformationalVersion"
    Private Const ASSEMBLY_FILE_VERSION As String = "AssemblyFileVersion"

    Private Const ASSEMBLY_VERSION_ATTRIBUTE As String = "AssemblyVersionAttribute"
    Private Const ASSEMBLY_INFORMATIONAL_VERSION_ATTRIBUTE As String = "AssemblyInformationalVersionAttribute"
    Private Const ASSEMBLY_FILE_VERSION_ATTRIBUTE As String = "AssemblyFileVersionAttribute"

#End Region

#Region "Enumerations"

    Public Enum Scale
      Major = 0
      Minor = 1
      Build = 2
      Revision = 3
    End Enum

    Public Enum VersionType
      Assembly = 0
      Product = 1
      File = 2
    End Enum

#End Region

#Region "Class Variables"

    Private mobjFileVersionInfo As FileVersionInfo

    Private mobjAssemblyVersion As New System.Version
    Private mobjProductVersion As New System.Version
    Private mobjFileVersion As New System.Version

#End Region

#Region "Public Properties"

    Public Property AssemblyVersion() As System.Version
      Get
        Return mobjAssemblyVersion
      End Get
      Set(ByVal value As System.Version)
        mobjAssemblyVersion = value
      End Set
    End Property

    Public Property ProductVersion() As System.Version
      Get
        Return mobjProductVersion
      End Get
      Set(ByVal value As System.Version)
        mobjProductVersion = value
      End Set
    End Property

    Public Property FileVersion() As System.Version
      Get
        Return mobjFileVersion
      End Get
      Set(ByVal value As System.Version)
        mobjFileVersion = value
      End Set
    End Property

    Public ReadOnly Property FileVersionInfo As FileVersionInfo
      Get
        Return mobjFileVersionInfo
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Shared Function Create(ByVal lpAssembly As Assembly) As AssemblyVersionInfo

      Dim lobjAssemblyVersionInfo As New AssemblyVersionInfo
      Dim lobjAttributeList As IList
      Dim lobjReflectedAssembly As Assembly = Nothing

      Try

        lobjAssemblyVersionInfo.AssemblyVersion = lpAssembly.GetName.Version



        Try
          lobjAttributeList = CustomAttributeData.GetCustomAttributes(lpAssembly)
        Catch FileLoadEx As FileLoadException
          ' Assembly.ReflectionOnlyLoad(FileLoadEx.FileName)
          lobjReflectedAssembly = Assembly.ReflectionOnlyLoad(lpAssembly.FullName)
          lobjAttributeList = CustomAttributeData.GetCustomAttributes(lobjReflectedAssembly)
        End Try


        For Each lobjCustomAttribute As CustomAttributeData In lobjAttributeList

          Debug.Print(String.Format("AttributeName: {0}", lobjCustomAttribute.Constructor.DeclaringType.Name))

          Select Case lobjCustomAttribute.Constructor.DeclaringType.Name
            Case ASSEMBLY_VERSION_ATTRIBUTE
              lobjAssemblyVersionInfo.AssemblyVersion = New Version(CType(lobjCustomAttribute.ConstructorArguments(0).Value, String))
            Case ASSEMBLY_INFORMATIONAL_VERSION_ATTRIBUTE
              lobjAssemblyVersionInfo.ProductVersion = New Version(CType(lobjCustomAttribute.ConstructorArguments(0).Value, String))
            Case ASSEMBLY_FILE_VERSION_ATTRIBUTE
              lobjAssemblyVersionInfo.FileVersion = New Version(CType(lobjCustomAttribute.ConstructorArguments(0).Value, String))
          End Select
        Next

        'If lobjAssemblyVersionInfo.AssemblyVersion Is Nothing Then
        '  lobjAssemblyVersionInfo.AssemblyVersion = lobjAssemblyVersionInfo.ProductVersion
        'End If

        Return lobjAssemblyVersionInfo

      Catch ex As Exception
        'ApplicationLogging.WriteLogEntry(ex.Message, Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 346356)
        ' Return what we can
        'Return lobjAssemblyVersionInfo
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Function

    'Public Shared Function Create(ByVal lpAssemblyInfo As AssemblyInfo) As AssemblyVersionInfo
    '  Try

    '    AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

    '    Dim lstrAssemblyPath As String = String.Format("{0}\{1}.exe", lpAssemblyInfo.DirectoryPath, lpAssemblyInfo.AssemblyName)
    '    Dim lobjAssembly As Assembly
    '    Dim lobjAssemblyVersionInfo As AssemblyVersionInfo

    '    If File.Exists(lstrAssemblyPath) = False Then
    '      ' It could not find the exe, let's try it as a dll
    '      lstrAssemblyPath = lstrAssemblyPath.Replace(".exe", ".dll")
    '      If File.Exists(lstrAssemblyPath) = False Then
    '        ' It can't find it as a dll either, throw an exception
    '        Throw New FileNotFoundException(String.Format(
    '          "Unable to create new AssemblyVersionInfo object using the specified AssemblyInfo parameter. " &
    '          " The assembly file '{0}' could not be found in the path '{1}'",
    '          Path.GetFileName(lstrAssemblyPath),
    '          Path.GetDirectoryName(lstrAssemblyPath)),
    '          lstrAssemblyPath)
    '      End If
    '    End If

    '    'AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

    '    lobjAssembly = Assembly.ReflectionOnlyLoadFrom(lstrAssemblyPath)

    '    lobjAssemblyVersionInfo = AssemblyVersionInfo.Create(lobjAssembly)

    '    lobjAssembly = Nothing

    '    GC.Collect()

    '    Return lobjAssemblyVersionInfo

    '  Catch ex As Exception
    '    'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Shared Function Create(ByVal lpAssemblyInfoFilePath As String) As AssemblyVersionInfo
      Try

        ' Make sure we got a valid info file path
        Helper.VerifyFilePath(lpAssemblyInfoFilePath, "lpAssemblyInfoFilePath", True)

        Dim lobjAssemblyVersion As System.Version = Nothing
        Dim lobjProductVersion As System.Version = Nothing
        Dim lobjFileVersion As System.Version = Nothing

        ' Open the file and read the contents into a string
        Dim lobjStreamReader As StreamReader = File.OpenText(lpAssemblyInfoFilePath)
        Dim lstrContents As String = lobjStreamReader.ReadToEnd

        ' Close the file
        lobjStreamReader.Close()

        ' Get the assembly version
        lobjAssemblyVersion = GetQualifiedVersion(lstrContents, VersionType.Assembly)

        ' Get the product version
        lobjProductVersion = GetQualifiedVersion(lstrContents, VersionType.Product)

        ' Get the file version
        lobjFileVersion = GetQualifiedVersion(lstrContents, VersionType.File)

        ' Return the combined version info
        Return New AssemblyVersionInfo(lobjAssemblyVersion, lobjProductVersion, lobjFileVersion)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetQualifiedVersion(ByVal lpAssemblyInfoFileContents As String, ByVal lpVersionType As VersionType) As System.Version
      Try
        Dim lobjRegex As Regex = Nothing
        Dim lobjMatches As MatchCollection = Nothing
        Dim lstrVersionNumber As String = "0.0.0.0"

        lobjRegex = New Regex(CreateVersionRegularExpression(lpVersionType))
        lobjMatches = lobjRegex.Matches(lpAssemblyInfoFileContents)

        If lobjMatches IsNot Nothing AndAlso lobjMatches.Count > 0 Then
          For Each lobjMatch As Match In lobjMatches
            If lobjMatch.Value.Contains("*") = False Then
              lstrVersionNumber = lobjMatch.Value
            End If
          Next
        End If

        Return New System.Version(lstrVersionNumber)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function IncrementAssemblyVersionInfo(ByVal lpAssemblyVersionInfo As AssemblyVersionInfo,
                                               ByVal lpScale As Scale) As AssemblyVersionInfo

      Try

        Dim lobjReturnAssemblyVersionInfo As AssemblyVersionInfo = lpAssemblyVersionInfo

        With lobjReturnAssemblyVersionInfo
          If lpScale < Scale.Build Then
            .AssemblyVersion = IncrementVersion(.AssemblyVersion, lpScale)
          End If
          .ProductVersion = IncrementVersion(.ProductVersion, lpScale)
          .FileVersion = IncrementVersion(.FileVersion, lpScale)
        End With

        Return lobjReturnAssemblyVersionInfo

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function IncrementVersion(ByVal lpVersionType As VersionType,
                                   ByVal lpScale As Scale) As System.Version
      Try
        Select Case lpVersionType
          Case VersionType.Assembly
            Return IncrementVersion(AssemblyVersion, lpScale)
          Case VersionType.Product
            Return IncrementVersion(ProductVersion, lpScale)
          Case VersionType.File
            Return IncrementVersion(FileVersion, lpScale)
            'Case VersionType.All
            '  IncrementVersion(AssemblyVersion, lpScale)
            '  IncrementVersion(ProductVersion, lpScale)
            '  Return IncrementVersion(FileVersion, lpScale)
          Case Else
            Throw New ArgumentOutOfRangeException
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function IncrementVersion(ByVal lpVersion As System.Version,
                                     ByVal lpScale As Scale) As System.Version
      Try

        Dim lobjNewVersion As System.Version = lpVersion

        Select Case lpScale
          Case Scale.Major
            lobjNewVersion = New System.Version(lpVersion.Major + 1, 0, 0, 0)
          Case Scale.Minor
            lobjNewVersion = New System.Version(lpVersion.Major, lpVersion.Minor + 1, 0, 0)
          Case Scale.Build
            lobjNewVersion = New System.Version(lpVersion.Major, lpVersion.Minor, lpVersion.Build + 1, 0)
          Case Scale.Revision
            lobjNewVersion = New System.Version(lpVersion.Major, lpVersion.Minor, lpVersion.Build, lpVersion.Revision + 1)
        End Select

        Return lobjNewVersion

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub IncrementAssemblyInfoFile(ByVal lpAssemblyInfoFilePath As String, ByVal lpScale As Scale)
      Try

        ' Get the current assembly info
        Dim lobjCurrentAssemblyInfo As AssemblyVersionInfo = AssemblyVersionInfo.Create(lpAssemblyInfoFilePath)

        Dim lobjNewVersionInfo As AssemblyVersionInfo = IncrementAssemblyVersionInfo(lobjCurrentAssemblyInfo, lpScale)

        UpdateVersion(lpAssemblyInfoFilePath, lobjNewVersionInfo)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub UpdateVersion(ByVal lpAssemblyInfoFilePath As String, ByVal lpNewVersion As AssemblyVersionInfo)
      Try

        ' Make sure we got a valid info file path
        Helper.VerifyFilePath(lpAssemblyInfoFilePath, "lpAssemblyInfoFilePath", True)

        ' Get the current assembly info
        Dim lobjCurrentAssemblyInfo As AssemblyVersionInfo = AssemblyVersionInfo.Create(lpAssemblyInfoFilePath)

        If lobjCurrentAssemblyInfo.Equals(lpNewVersion) = True Then
          ' There is no work to do
          Exit Sub
        Else

          Dim lstrContents As String

          ' Open the file and read the contents into a string
          Using lobjStreamReader As StreamReader = File.OpenText(lpAssemblyInfoFilePath)
            lstrContents = lobjStreamReader.ReadToEnd
            ' Close the file
            lobjStreamReader.Close()
          End Using

          Dim lobjRegex As Regex = Nothing

          ' Replace the old version information with the new version information

          ' Replace the Assembly version
          lobjRegex = New Regex(CreateVersionRegularExpression(VersionType.Assembly))
          lstrContents = lobjRegex.Replace(lstrContents, lpNewVersion.AssemblyVersion.ToString)

          ' Replace the Product version
          lobjRegex = New Regex(CreateVersionRegularExpression(VersionType.Product))
          lstrContents = lobjRegex.Replace(lstrContents, lpNewVersion.ProductVersion.ToString)

          ' Replace the File version
          lobjRegex = New Regex(CreateVersionRegularExpression(VersionType.File))
          lstrContents = lobjRegex.Replace(lstrContents, lpNewVersion.FileVersion.ToString)

          ' Write the contents back to the file
          Using lobjStreamWriter As StreamWriter = New StreamWriter(lpAssemblyInfoFilePath, False, System.Text.Encoding.Unicode)
            lobjStreamWriter.Write(lstrContents)
            lobjStreamWriter.Close()
          End Using

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Shared Function CurrentDomain_ReflectionOnlyAssemblyResolve(ByVal sender As Object, ByVal args As ResolveEventArgs) As Assembly
      Try
        '' Search the assembly in a different directory.
        'Dim searchDir As String = "c:\myassemblies"
        'For Each dllFile As String In Directory.GetFiles(searchDir, "*.dll")
        '  Try
        '    Dim asm As Assembly = Assembly.LoadFile(dllFile)
        '    ' If the DLL is an assembly and its name matches, we've found it.
        '    If asm.GetName().Name = e.Name Then Return asm
        '  Catch ex As Exception
        '    ' Ignore DLLs that aren't valid assemblies.
        '  End Try
        'Next
        '' If we get here, return Nothing to signal that the search failed.
        'Return Nothing
        Return System.Reflection.Assembly.ReflectionOnlyLoad(args.Name)
      Catch FileNotFoundEx As FileNotFoundException
        If args.Name.ToLower.Contains("xmlserializers") Then
          ' Ignore for these
          Return Nothing
        Else
          'ApplicationLogging.WriteLogEntry(String.Format("Failed to load assembly '{0}' in reflection only mode: {1}", _
          '                                               args.Name, FileNotFoundEx.Message), TraceEventType.Verbose, 1011)
          Return Nothing
        End If
      Catch ex As Exception
        'ApplicationLogging.WriteLogEntry(String.Format("Failed to load assembly '{0}' in reflection only mode: {1}", _
        '                                               args.Name, ex.Message), TraceEventType.Verbose, 1011)
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Return nothing to signal that the search failed
        Return Nothing
      End Try
    End Function

    Private Shared Function VersionEntryString(ByVal lpAttributeName As String, ByVal lpVersion As System.Version) As String
      Try
        Return String.Format("{0}({1}{2}{1})", lpAttributeName, ControlChars.Quote, lpVersion.ToString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function CreateVersionRegularExpression(ByVal lpVersionType As VersionType) As String
      Try

        Dim lstrAttributeName As String = Nothing
        Dim lstrVersionType As String = Nothing

        Select Case lpVersionType
          Case VersionType.Assembly
            lstrVersionType = ASSEMBLY_VERSION
          Case VersionType.Product
            lstrVersionType = ASSEMBLY_INFORMATIONAL_VERSION
          Case VersionType.File
            lstrVersionType = ASSEMBLY_FILE_VERSION
          Case Else
            Throw New ArgumentOutOfRangeException("lpVersionType", lpVersionType, String.Format("The value '{0}' for VersionType is out of the expected range.", lpVersionType.ToString))
        End Select

        Return String.Format("(?<Version>(?<={0}\(\""|{0}Attribute\(\"").*(?=\""\)))", lstrVersionType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function DebuggerIdentifier() As String

      Dim lobjIdentifierBuilder As New StringBuilder

      Try

        With lobjIdentifierBuilder
          .AppendFormat("Assembly: {0}, ", AssemblyVersion.ToString)
          .AppendFormat("Product: {0}, ", ProductVersion.ToString)
          .AppendFormat("File: {0}", FileVersion.ToString)

          Return .ToString

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpAssemblyVersion As System.Version,
                 ByVal lpProductVersion As System.Version,
                 ByVal lpFileVersion As System.Version)

      AssemblyVersion = lpAssemblyVersion
      ProductVersion = lpProductVersion
      FileVersion = lpFileVersion
    End Sub

    Public Sub New(ByVal lpAssemblyVersionString As String,
               ByVal lpProduct_FileVersionString As String)

      Try
        AssemblyVersion = New System.Version(lpAssemblyVersionString)
        ProductVersion = New System.Version(lpProduct_FileVersionString)
        FileVersion = New System.Version(lpProduct_FileVersionString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Sub New(ByVal lpAssemblyVersionString As String,
                 ByVal lpProductVersionString As String,
                 ByVal lpFileVersionString As String)

      Try
        AssemblyVersion = New System.Version(lpAssemblyVersionString)
        ProductVersion = New System.Version(lpProductVersionString)
        FileVersion = New System.Version(lpFileVersionString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Sub New(ByVal lpAssembly As Assembly)
      Dim lobjAssemblyVersionInfo As AssemblyVersionInfo

      Try

        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

        lobjAssemblyVersionInfo = AssemblyVersionInfo.Create(lpAssembly)

        Helper.AssignObjectProperties(lobjAssemblyVersionInfo, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpAssemblyPath As String)
      Try
        Dim lobjAssembly As Assembly
        Dim lobjAssemblyVersionInfo As AssemblyVersionInfo

        lobjAssemblyVersionInfo = New AssemblyVersionInfo(FileVersionInfo.GetVersionInfo(lpAssemblyPath))

        Helper.AssignObjectProperties(lobjAssemblyVersionInfo, Me)

        lobjAssembly = Nothing

        GC.Collect()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFileVersionInfo As FileVersionInfo)
      Try
        mobjFileVersionInfo = lpFileVersionInfo

        With mobjFileVersionInfo
          mobjAssemblyVersion = New System.Version(.ProductMajorPart, .ProductMinorPart, 0, 0)
          mobjProductVersion = New System.Version(.ProductVersion)
          mobjFileVersion = New System.Version(.FileVersion)
        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Sub New(ByVal lpAssemblyInfo As AssemblyInfo)
    '  Try

    '    AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

    '    Dim lstrAssemblyPath As String = String.Format("{0}\{1}.exe", lpAssemblyInfo.DirectoryPath, lpAssemblyInfo.AssemblyName)
    '    Dim lobjAssembly As Assembly
    '    Dim lobjAssemblyVersionInfo As AssemblyVersionInfo

    '    If File.Exists(lstrAssemblyPath) = False Then
    '      ' It could not find the exe, let's try it as a dll
    '      lstrAssemblyPath = lstrAssemblyPath.Replace(".exe", ".dll")
    '      If File.Exists(lstrAssemblyPath) = False Then
    '        ' It can't find it as a dll either, throw an exception
    '        Throw New FileNotFoundException(String.Format(
    '          "Unable to create new AssemblyVersionInfo object using the specified AssemblyInfo parameter. " &
    '          " The assembly file '{0}' could not be found in the path '{1}'",
    '          Path.GetFileName(lstrAssemblyPath),
    '          Path.GetDirectoryName(lstrAssemblyPath)),
    '          lstrAssemblyPath)
    '      End If
    '    End If

    '    'AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

    '    lobjAssembly = Assembly.ReflectionOnlyLoadFrom(lstrAssemblyPath)

    '    lobjAssemblyVersionInfo = AssemblyVersionInfo.Create(lobjAssembly)

    '    Helper.AssignObjectProperties(lobjAssemblyVersionInfo, Me)

    '    lobjAssembly = Nothing

    '    GC.Collect()

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Equality Implementations"

    Public Shadows Function Equals(ByVal x As Object, ByVal y As Object) As Boolean Implements System.Collections.IEqualityComparer.Equals
      Try
        If TypeOf (x) Is AssemblyVersionInfo AndAlso TypeOf (y) Is AssemblyVersionInfo Then
          Return DirectCast(x, AssemblyVersionInfo).Equals(DirectCast(y, AssemblyVersionInfo))
        Else
          If TypeOf (x) Is AssemblyVersionInfo = False Then
            Throw New ArgumentException(String.Format("Object x is of type '{0}', type AssemblyVersionInfo was expected", x.GetType.Name, "x"))
          End If

          If TypeOf (y) Is AssemblyVersionInfo = False Then
            Throw New ArgumentException(String.Format("Object y is of type '{0}', type AssemblyVersionInfo was expected", y.GetType.Name, "y"))
          End If

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function GetHashCode(ByVal obj As Object) As Integer Implements System.Collections.IEqualityComparer.GetHashCode
      Return MyBase.GetHashCode
    End Function

    Public Shadows Function Equals(ByVal other As AssemblyVersionInfo) As Boolean Implements System.IEquatable(Of AssemblyVersionInfo).Equals
      Try
        If AssemblyVersion.Equals(other.AssemblyVersion) AndAlso
        ProductVersion.Equals(other.ProductVersion) AndAlso
        FileVersion.Equals(other.FileVersion) Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace