'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Collection of Version objects.</summary>
  <Serializable()>
  Partial Public Class Versions
    Inherits CCollection(Of Version)
    Implements ICloneable

#Region "Public Methods"

    Public Sub SortProperties()
      Try
        For Each lobjECMVersion As Version In Me
          lobjECMVersion.Properties.Sort()
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Versions::SortProperties")
      End Try
    End Sub

#End Region

#Region "ICloneable Support"

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjVersions As New Versions()

      Try
        For Each lobjVersion As Version In Me
          lobjVersions.Add(lobjVersion.Clone)
        Next
        Return lobjVersions
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace