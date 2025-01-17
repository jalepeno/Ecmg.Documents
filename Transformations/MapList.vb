'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>
  ''' Manages a set of ValueMaps for performing simple data lookups
  ''' </summary>
  ''' <remarks>The original values must be unique</remarks>
  <Serializable()>
  Public Class MapList
    Inherits ObservableCollection(Of ValueMap)

#Region "Class Variables"

    'Private lblnCaseSensitive As Boolean = True
    Private mobjDataList As DataList

#End Region

#Region "Public Properties"

    '<XmlAttribute("caseSensitive", GetType(Boolean))> _
    'Public Property CaseSensitive() As Boolean
    '  Get
    '    Return lblnCaseSensitive
    '  End Get
    '  Set(ByVal value As Boolean)
    '    lblnCaseSensitive = value
    '  End Set
    'End Property

    Public ReadOnly Property DataList() As DataList
      Get
        Return mobjDataList
      End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Searches through the set of value maps looking for the 
    ''' specified original and returns the corresponding replacement
    ''' </summary>
    ''' <param name="original"></param>
    ''' <returns>The matching replacement if the original is found, otherwise returns the original</returns>
    ''' <remarks></remarks>
    Public Function FindReplacement(ByVal original As String) As String

      Try

        ' Determine whether or not the search will be case sensitive
        ' Depending on case sensitivity we will use a different compare method

        If DataList.CaseSensitive Then
          ' The search is case sensitive

          ' Loop through each value map
          For Each lobjValueMap As ValueMap In Me

            If StrComp(original, lobjValueMap.Original, CompareMethod.Binary) = 0 Then
              Return lobjValueMap.Replacement
            End If

          Next

          Return original

        Else
          ' The search is not case sensitive

          ' Loop through each value map
          For Each lobjValueMap As ValueMap In Me

            If StrComp(original, lobjValueMap.Original, CompareMethod.Text) = 0 Then
              Return lobjValueMap.Replacement
            End If

          Next

          'ApplicationLogging.WriteLogEntry(String.Format("No replacement found for original value of '{0}'.", original), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 63404)
          Return original
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Adds a new ValueMap using the specified orginal and replacement values
    ''' </summary>
    ''' <param name="original">The Original property for the ValueMap</param>
    ''' <param name="replacement">The Replacement property for the ValueMap</param>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">If either the original or replacement values are null an ArgumentNullException is thrown.  Empty Strings are acceptable.</exception>
    Public Overloads Sub Add(ByVal original As String, ByVal replacement As String)
      Try

        ' Make sure we get a valid string
        If original Is Nothing Then
          Throw New ArgumentNullException("original")
        End If

        ' Make sure we get a valid string
        If replacement Is Nothing Then
          Throw New ArgumentNullException("replacement")
        End If

        Add(New ValueMap(original, replacement))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal map As ValueMap)
      Try

        ' Make sure the value for original will be unique
        Dim lobjMap As New ValueMap
        If OriginalExists(map.Original, lobjMap) Then
          Throw New OriginalNotUniqueException(lobjMap, map.Original)
        End If

        MyBase.Add(map)

      Catch NotUniqueEx As OriginalNotUniqueException
        If Helper.IsDeserializationBasedCall Then

          ' Make sure this is not a phantom coming from deserialization
          If map.Original Is Nothing AndAlso map.Replacement Is Nothing Then
            ' This is an empty map, just ignore completely
            Exit Sub
          End If
          ' If this is a result of deserialization then just log it and keep going
          ApplicationLogging.WriteLogEntry(String.Format("Unable to add map: {0}", NotUniqueEx.Message),
            TraceEventType.Warning, 888)
        Else
          ApplicationLogging.LogException(NotUniqueEx, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Checks to see if the specified original value exists in the list
    ''' </summary>
    ''' <param name="original">The original value to test for</param>
    ''' <returns>True if the value exists and false if it does not</returns>
    ''' <remarks>The test may or may not be case sensitive 
    ''' depending on the value of the CaseSensitive property</remarks>
    Public Function OriginalExists(ByVal original As String) As Boolean
      Try

        ' Declare a value map the pass, even though we will ignore the value
        Dim lobjValueMap As New ValueMap

        Return OriginalExists(original, lobjValueMap)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the specified original value exists in the list
    ''' </summary>
    ''' <param name="original">The original value to test for</param>
    ''' <param name="existingValueMap">Returns a reference to the existing ValueMap 
    ''' containing the specified original value</param>
    ''' <returns>True if the value exists and false if it does not</returns>
    ''' <remarks>The test may or may not be case sensitive 
    ''' depending on the value of the CaseSensitive property</remarks>
    Public Function OriginalExists(ByVal original As String, ByRef existingValueMap As ValueMap) As Boolean
      Try

        ' Determine whether or not the search will be case sensitive
        ' Depending on case sensitivity we will use a different compare method

        If DataList.CaseSensitive Then
          ' The search is case sensitive

          ' Loop through each value map
          For Each lobjValueMap As ValueMap In Me
            If StrComp(original, lobjValueMap.Original, CompareMethod.Binary) = 0 Then
              existingValueMap = lobjValueMap
              Return True
            End If
          Next
          Return False
        Else
          ' The search is not case sensitive

          ' Loop through each value map
          For Each lobjValueMap As ValueMap In Me
            If StrComp(original, lobjValueMap.Original, CompareMethod.Text) = 0 Then
              existingValueMap = lobjValueMap
              Return True
            End If
          Next
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetMapListItems(lpDelimiter As String) As ObservableCollection(Of String)
      Try
        Dim lobjStringList As New ObservableCollection(Of String)

        For Each lobjValueMap As ValueMap In Me
          lobjStringList.Add(String.Format("{0}{1}{2}", lobjValueMap.Original, lpDelimiter, lobjValueMap.Replacement))
        Next

        Return lobjStringList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Public Sub New()

    End Sub

    Public Sub New(ByVal dataList As DataList)
      mobjDataList = dataList
    End Sub

  End Class

  ''' <summary>
  ''' Thrown when an attempt to add a non-unique original occurs
  ''' </summary>
  ''' <remarks></remarks>
  Public Class OriginalNotUniqueException
    Inherits ApplicationException

#Region "Class Variables"

    Private mobjExistingValueMap As ValueMap

#End Region

#Region "Public Properties"

    Public Property ExistingValueMap() As ValueMap
      Get
        Return mobjExistingValueMap
      End Get
      Set(ByVal value As ValueMap)
        Try

          ' Make sure we get a valid ValueMap
          If value Is Nothing Then
            Throw New ArgumentNullException()
          End If

          mobjExistingValueMap = value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal existingOriginal As String)
      MyBase.New(String.Format("The original value of '{0}' is not unique in the list", existingOriginal))
    End Sub

    Public Sub New(ByVal existingValueMap As ValueMap, ByVal existingOriginal As String)

      MyBase.New(String.Format("The original value of '{0}' is not unique in the list", existingOriginal))

      Me.ExistingValueMap = existingValueMap

    End Sub

    Public Sub New(ByVal message As String, ByVal existingValueMap As ValueMap)

      MyBase.New(message)

      Try

        ' Make sure we get a valid ValueMap
        If existingValueMap Is Nothing Then
          Throw New ArgumentNullException("existingValueMap")
        End If

        Me.ExistingValueMap = existingValueMap

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
