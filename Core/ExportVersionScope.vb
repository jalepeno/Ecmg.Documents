'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>
  ''' Used to control which 
  ''' versions of a document are to 
  ''' be exported.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ExportVersionScope

#Region "Class Variables"

    Private menuVersionScope As VersionScopeEnum = VersionScopeEnum.AllVersions
    Private mintMaxVersions As Integer

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the scope of which versions to export.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Scope As VersionScopeEnum
      Get
        Return menuVersionScope
      End Get
      Set(value As VersionScopeEnum)
        menuVersionScope = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the maximum 
    ''' number of versions to export.
    ''' </summary>
    ''' <value>
    ''' An integer representing the maximum 
    ''' number of versions to export.  A value 
    ''' of zero indicates there is no maximum 
    ''' and all available versions for the 
    ''' specified scope will be exported.
    ''' </value>
    ''' <returns></returns>
    ''' <remarks>This setting only applies for Scope values 
    ''' of LastNVersions and FirstNVersions.</remarks>
    Public Property MaxVersions As Integer
      Get
        Return mintMaxVersions
      End Get
      Set(value As Integer)
        mintMaxVersions = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Constructs a new ExportVersionScope object with the specified parameters.
    ''' </summary>
    ''' <param name="lpScope">
    ''' The version scope for the export.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(lpScope As VersionScopeEnum)
      Try
        Scope = lpScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Constructs a new ExportVersionScope object with the specified parameters.
    ''' </summary>
    ''' <param name="lpScope">
    ''' The version scope for the export.
    ''' </param>
    ''' <param name="lpMaxVersions">
    ''' The maximum number of versions to export 
    ''' (if applicable for the specified scope).
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(lpScope As VersionScopeEnum, lpMaxVersions As Integer)
      Try
        Scope = lpScope
        MaxVersions = lpMaxVersions
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
