'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Providers

  Public Class ActionProperties
    Inherits CCollection(Of ActionProperty)

#Region "Overriden Methods"

    Public Overloads Function Contains(ByVal lpProperty As IProperty) As Boolean
      Try

        If lpProperty Is Nothing Then
          Return False
        End If

        For Each lobjProperty As ActionProperty In Me
          If String.Equals(lobjProperty.Name, lpProperty.Name, StringComparison.InvariantCultureIgnoreCase) Then
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
    ''' Locates and returns the property by name
    ''' </summary>
    ''' <param name="lpName">The name of the property</param>
    ''' <returns>The property if found, otherwise returns nothing</returns>
    ''' <remarks></remarks>
    Default Public Shadows ReadOnly Property Item(ByVal lpName As String) As ActionProperty
      Get
        Try
          For Each lobjProperty As ActionProperty In Me
            If String.Equals(lobjProperty.Name, lpName, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjProperty
            End If
          Next

          Return Nothing

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get

    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Returns a string array of all the ActionProperty names in the collection
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NameArray() As String()
      Try
        Dim lstrNameArray As String()
        ReDim lstrNameArray(Me.Count - 1)
        Dim lintArrayCounter As Integer

        For Each lobjProperty As ActionProperty In Me
          lstrNameArray(lintArrayCounter) = lobjProperty.Name
          lintArrayCounter += 1
        Next

        Return lstrNameArray

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
