'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities
Imports Newtonsoft.Json


#End Region

Namespace Core
  ''' <summary>Fully describes a single document version.</summary>
  <Serializable()>
  Partial Public Class Version
    Implements IComparable
    Implements ICloneable

    ' TODO: Add ToDataTable method to return all the version information in a data table

#Region "Class Variables"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjFileNames As New List(Of String)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the names of all the files for all the content elements.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property FileNames As IList(Of String)
      Get
        Try
          Return mobjFileNames
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Functions"

    ''' <summary>
    ''' Changes the value of the specifed property in the document
    ''' </summary>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>If the property scope is set to the version level, only the version specified by the lpVersionIndex will be affected.</remarks>
    Public Function ChangePropertyValue(ByVal lpName As String,
                                        ByVal lpNewValue As Object) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Version::ChangePropertyValue", TraceEventType.Verbose)

      Try

        Try


          ' Get the specified property first
          Dim lobjProperty As ECMProperty = Properties(lpName)

          If lobjProperty Is Nothing Then
            ApplicationLogging.WriteLogEntry(
              String.Format(
                "Unable to change the value of property '{0}' to '{1}' in Version::ChangePropertyValue, the property does not exist in this version",
                lpName, lpNewValue),
              TraceEventType.Warning, 4582)

            Return False

          Else
            lobjProperty.ChangePropertyValue(lpNewValue)
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ApplicationLogging.WriteLogEntry("Exit Version::ChangePropertyValue", TraceEventType.Verbose)
          Throw New InvalidOperationException("Could not change property value [" & ex.Message & "]", ex)
        End Try

        ApplicationLogging.WriteLogEntry("Exit Version::ChangePropertyValue", TraceEventType.Verbose)

        Return True
        ' Change the value of the specified property

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Sub ClearPropertyValue(ByVal lpName As String)
      Try
        Try

          ' Get the specified property first
          Dim lobjProperty As ECMProperty = Properties(lpName)

          If lobjProperty Is Nothing Then
            ApplicationLogging.WriteLogEntry(
              String.Format(
                "Unable to clear the value of property '{0}' in Version::ClearPropertyValue, the property does not exist in this version",
                lpName),
              TraceEventType.Warning, 4585)

          Else
            lobjProperty.ClearPropertyValue()
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ApplicationLogging.WriteLogEntry("Exit Version::ClearPropertyValue", TraceEventType.Verbose)
          Throw New InvalidOperationException("Could not clear property value [" & ex.Message & "]", ex)
        End Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function GetLargestFileSize() As FileSize
      Dim lobjLargestFileSize As FileSize = New FileSize(0)
      For Each lobjContent As Content In Contents
        If lobjContent.FileSize.CompareTo(lobjLargestFileSize) > 0 Then
          lobjLargestFileSize = lobjContent.FileSize
        End If
      Next

      Return lobjLargestFileSize

    End Function

    Public Function SetPropertyValue(ByVal lpPropertyName As String,
                                 ByVal lpPropertyValue As Object, ByVal lpCreateProperty As Boolean,
                                 Optional ByVal lpPropertyType As PropertyType = PropertyType.ecmString) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Version::SetPropertyValue", TraceEventType.Verbose)
      Try

        If Properties.PropertyExists(lpPropertyName) = False Then
          ' The proeprty does not currently exist
          If lpCreateProperty = True Then
            'Properties.Add(New ECMProperty(lpPropertyType, lpPropertyName, lpPropertyValue))
            Properties.Add(PropertyFactory.Create(lpPropertyType, lpPropertyName, lpPropertyValue))
            Return True
          End If
          Throw New InvalidOperationException(String.Format(
                                                            "Cannot set the value of '{0}' to '{1}' as the property does not exist in the document and the parameter lpCreateProperty was set to False.",
                                                            lpPropertyValue.ToString, lpPropertyName))
        Else

          Properties(lpPropertyName).Value = lpPropertyValue
          Return True

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ApplicationLogging.WriteLogEntry("Exit Version::SetPropertyValue", TraceEventType.Verbose)
      End Try
    End Function

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Just revert to the default behavior
        Return Me.GetType.Name
      End Try
    End Function

    Public Function RemoveFromDocument() As Boolean
      Try

        If Document Is Nothing Then
          Throw New Exceptions.DocumentReferenceNotSetException("Unable to remove version, the Document reference is not set.")
        End If

        If Document.Versions.Contains(Me) Then
          Document.Versions.Remove(Me)
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

#End Region

#Region "Private Methods"

    Private Sub mobjContents_CollectionChanged(sender As Object, e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles MobjContents.CollectionChanged
      Try
        Select Case e.Action

          Case Specialized.NotifyCollectionChangedAction.Add
            ' Add the new names that are not already in the list.
            For Each lobjContent As Content In e.NewItems
              If ((Not String.IsNullOrEmpty(lobjContent.FileName)) AndAlso
                  (Not Me.mobjFileNames.Contains(lobjContent.FileName))) Then
                mobjFileNames.Add(lobjContent.FileName)
              End If
            Next

          Case Specialized.NotifyCollectionChangedAction.Remove
            ' Remove the old names that might be in the list.
            For Each lobjContent As Content In e.OldItems
              If ((Not String.IsNullOrEmpty(lobjContent.FileName)) AndAlso
                  (Me.mobjFileNames.Contains(lobjContent.FileName))) Then
                mobjFileNames.Remove(lobjContent.FileName)
              End If
            Next

          Case Specialized.NotifyCollectionChangedAction.Replace
            ' Remove the old names that might be in the list.
            For Each lobjContent As Content In e.OldItems
              If ((Not String.IsNullOrEmpty(lobjContent.FileName)) AndAlso
                  (Me.mobjFileNames.Contains(lobjContent.FileName))) Then
                mobjFileNames.Remove(lobjContent.FileName)
              End If
            Next
            ' Add the new names that are not already in the list.
            For Each lobjContent As Content In e.NewItems
              If ((Not String.IsNullOrEmpty(lobjContent.FileName)) AndAlso
                  (Not Me.mobjFileNames.Contains(lobjContent.FileName))) Then
                mobjFileNames.Add(lobjContent.FileName)
              End If
            Next

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      If TypeOf obj Is Version Then
        Return ID.CompareTo(obj.ID)
      Else
        Throw New ArgumentException("Object is not a Version")
      End If

    End Function

#End Region

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjVersion As New Version()
      Try
        lobjVersion.Properties = Me.Properties.Clone
        lobjVersion.Contents = Me.Contents.Clone
        Return lobjVersion
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

  End Class

End Namespace