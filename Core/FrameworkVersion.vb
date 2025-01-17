'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class FrameworkVersion

#Region "Class Variables"

    Friend Shared mobjAssemblyVersion As AssemblyVersionInfo 'New AssemblyVersionInfo(Assembly.GetAssembly(Me.GetType))
    Private Shared mobjInstance As New FrameworkVersion

#End Region

#Region "Public Properties"

    Public Shared ReadOnly Property AssemblyVersion() As AssemblyVersionInfo
      Get
        Return mobjAssemblyVersion
      End Get
    End Property

    Public Shared ReadOnly Property Instance() As FrameworkVersion
      Get
        If mobjInstance Is Nothing Then
          mobjInstance = New FrameworkVersion
        End If
        Return mobjInstance
      End Get
    End Property

    ''' <summary>
    ''' Gets a string representation of the current 
    ''' executing file version of the Ecmg.Cts.dll.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property CurrentVersion() As String
      Get
        Return AssemblyVersion.FileVersion.ToString
      End Get
    End Property

#End Region

#Region "Constructors"

    Private Sub New()
      mobjAssemblyVersion = New AssemblyVersionInfo(Assembly.GetAssembly(Me.GetType))
    End Sub

#End Region

#Region "Public Methods"

    Public Shared Shadows Function ToString() As String
      Return AssemblyVersion.FileVersion.ToString
    End Function

#End Region

#Region "Private Methods"
    Private Sub Initialize()
      mobjInstance = New FrameworkVersion
    End Sub
#End Region

  End Class

End Namespace
