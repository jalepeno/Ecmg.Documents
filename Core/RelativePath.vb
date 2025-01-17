'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core

  ''' <summary>
  ''' This class provides functions used to generate a relative path
  ''' from two absolute paths.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class RelativePath

    ''' <summary>
    ''' Break a path down into individual elements and add to a list.
    ''' Example: If a path is /a/b/c/d.txt, 
    ''' the breakdown will be (d.txt, c, b, a)
    ''' </summary>
    ''' <param name="lpFilePath"></param>
    ''' <returns>A List(Of String) with the individual elements of the path in reverse order</returns>
    ''' <remarks></remarks>
    Private Shared Function GetPathList(ByVal lpFilePath As String,
                                        ByVal lpFileSeparator As String) As List(Of String)

      Dim lobjList As New List(Of String)
      Dim lstrArray() As String
      'Dim lobjFileInfo As FileInfo
      'Dim lobjDirectoryInfo As DirectoryInfo
      'Dim lobjParentDirectory As DirectoryInfo

      Try
        'If File.Exists(lpFilePath) Then
        '  lobjFileInfo = New FileInfo(lpFilePath)
        '  lobjList.Add(lobjFileInfo.Name)
        '  ' Add the file name
        '  lobjDirectoryInfo = New DirectoryInfo(lobjFileInfo.DirectoryName)
        'ElseIf Directory.Exists(lpFilePath) Then
        '  lobjDirectoryInfo = New DirectoryInfo(lpFilePath)
        'Else
        '  Throw New FileNotFoundException(String.Format("The path '{0}' is not a valid file or directory path", lpFilePath), lpFilePath)
        'End If


        ' Add the current directory 
        'lobjList.Add(lobjDirectoryInfo.Name)

        'lobjParentDirectory = lobjDirectoryInfo.Parent

        'Do Until lobjParentDirectory Is Nothing
        '  lobjList.Add(lobjParentDirectory.Name)
        '  lobjParentDirectory = lobjParentDirectory.Parent
        'Loop

        lstrArray = lpFilePath.Split(lpFileSeparator)
        For i As Integer = lstrArray.Length - 1 To 0 Step -1
          lobjList.Add(lstrArray(i))
        Next

        Return lobjList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("RelativePath::GetPathList('{0}')", lpFilePath))
        Throw New Exception(String.Format("Unable to GetPathList: {0}", ex.Message), ex)
      End Try

    End Function

    ''' <summary>
    ''' Figure out a string representing the relative path of 'lpFilePathList' with respect to 'lpHomePathList'
    ''' </summary>
    ''' <param name="lpHomePathList"></param>
    ''' <param name="lpFilePathList"></param>
    ''' <param name="lpFileSeparator">Typically either \ or /</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function MatchPathLists(ByVal lpHomePathList As List(Of String),
                                           ByVal lpFilePathList As List(Of String),
                                           ByVal lpFileSeparator As String) As String

      Dim lintHomePathLevelCounter As Integer
      Dim lintFilePathLevelCounter As Integer
      Dim lstrRelativePath As String = ""

      Try
        'ApplicationLogging.WriteLogEntry("Enter RelativePath::MatchPathLists", TraceEventType.Verbose)

        ' Start at the beginning of the of the lists
        ' Iterate while both lists are equal
        lintHomePathLevelCounter = lpHomePathList.Count - 1
        lintFilePathLevelCounter = lpFilePathList.Count - 1

        ' First eliminate the common root
        While (lintHomePathLevelCounter >= 1) AndAlso
              (lpHomePathList(lintHomePathLevelCounter).Equals(lpFilePathList(lintFilePathLevelCounter)))

          lintHomePathLevelCounter -= 1
          lintFilePathLevelCounter -= 1

        End While

        If (lintHomePathLevelCounter = 0 AndAlso lintFilePathLevelCounter = 0) AndAlso (lpFilePathList(1) = lpHomePathList(1)) Then
          lstrRelativePath = String.Format("..\{0}", lpFilePathList(0))
          Return lstrRelativePath
        End If

        ' For each remaining level in the home path, add a ..
        For lintCounter As Integer = lintHomePathLevelCounter - 1 To 1 Step -1
          lstrRelativePath += String.Format("..{0}", lpFileSeparator)
        Next


        ' For each level in the file path, add the path
        For lintFPCounter As Integer = lintFilePathLevelCounter - 1 To 1 'Step -1
          lstrRelativePath += String.Format("{0}{1}",
                                            lpFilePathList(lintFPCounter), lpFileSeparator)
        Next

        ' Add the file name
        lstrRelativePath += lpFilePathList(0)

        Return lstrRelativePath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "RelativePath::MatchPathLists")
        ApplicationLogging.WriteLogEntry("Since there was an exception encountered we will return an empty string.", TraceEventType.Information)
        'ApplicationLogging.WriteLogEntry("Exit RelativePath::MatchPathLists", TraceEventType.Verbose)
        Return ""
      Finally
        'ApplicationLogging.WriteLogEntry("Exit RelativePath::MatchPathLists", TraceEventType.Verbose)
      End Try

    End Function

    ''' <summary>
    ''' Get the relative path of string 'lpFilePath' with respect
    ''' to 'lpHomeDirectory' directory
    ''' Example:  lpHomeDirectory = /a/b/c
    '''           lpFilePath = /a/d/e/x.txt
    '''           RelativePath = GetRelativePath(lpHomeDirectory, lpFilePath) = ../../d/e/x.txt
    ''' </summary>  
    ''' <param name="lpHomeDirectory">Should be a directory, not a file, or it doesn't make sense.</param>
    ''' <param name="lpFilePath">The file to generate a relative path for.</param>
    ''' <returns>Relative path from 'lpHomeDirectory' to 'lpFilePath' as a string.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetRelativePath(ByVal lpHomeDirectory As String,
                                           ByVal lpFilePath As String) As String
      Try
        Return GetRelativePath(lpHomeDirectory, lpFilePath, IO.Path.PathSeparator)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Get the relative path of string 'lpFilePath' with respect
    ''' to 'lpHomeDirectory' directory
    ''' Example:  lpHomeDirectory = /a/b/c
    '''           lpFilePath = /a/d/e/x.txt
    '''           RelativePath = GetRelativePath(lpHomeDirectory, lpFilePath) = ../../d/e/x.txt
    ''' </summary>  
    ''' <param name="lpHomeDirectory">Should be a directory, not a file, or it doesn't make sense.</param>
    ''' <param name="lpFilePath">The file to generate a relative path for.</param>
    ''' <param name="lpFileSeparator">Typically either \ or /</param>
    ''' <returns>Relative path from 'lpHomeDirectory' to 'lpFilePath' as a string.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetRelativePath(ByVal lpHomeDirectory As String,
                                           ByVal lpFilePath As String,
                                           ByVal lpFileSeparator As String) As String

      Dim lobjHomePathList As List(Of String)
      Dim lobjFilePathList As List(Of String)
      Dim lstrRelativePath As String
      Dim lstrHomeDirectory As String

      If IO.File.Exists(lpHomeDirectory) Then
        lstrHomeDirectory = IO.Path.GetDirectoryName(lpHomeDirectory)
      ElseIf IO.Directory.Exists(lpHomeDirectory) Then
        lstrHomeDirectory = lpHomeDirectory
      ElseIf IO.Directory.Exists(IO.Path.GetDirectoryName(lpHomeDirectory)) Then
        lstrHomeDirectory = IO.Path.GetDirectoryName(lpHomeDirectory)
      Else
        Throw New Exceptions.InvalidPathException(lpHomeDirectory)
      End If

      Try
        'ApplicationLogging.WriteLogEntry("Enter RelativePath::GetRelativePath", TraceEventType.Verbose)
        lobjHomePathList = GetPathList(lstrHomeDirectory, lpFileSeparator)
        lobjFilePathList = GetPathList(lpFilePath, lpFileSeparator)

        lstrRelativePath = MatchPathLists(lobjHomePathList, lobjFilePathList, lpFileSeparator)

        'ApplicationLogging.WriteLogEntry("Exit RelativePath::GetRelativePath", TraceEventType.Verbose)
        Return lstrRelativePath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("RelativePath::GetRelativePath('{0}', '{1}', '{2}')", lpHomeDirectory, lpFilePath, lpFileSeparator))
        Throw New Exception("Unable to get relative path", ex)
      End Try

    End Function

    'Public Shared Function GetCurrentPath(ByVal lpCDFPath As String, ByVal lpRelativeContentPath As String) As String
    '  Try

    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

  End Class
End Namespace