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

  ''' <summary>
  ''' Represents a special property used not for setting an index 
  ''' value but for setting some other attribute of a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ActionProperty
    Inherits ECMProperty

#Region "Constructors"

    Public Sub New()
      MyBase.New(Nothing)
    End Sub

    Public Sub New(ByVal lpName As String,
                   ByVal lpDefaultValue As String,
                   ByVal lpType As PropertyType,
                   ByVal lpDescription As String)

      MyBase.New(Nothing)

      Try
        Type = lpType
        Name = lpName
        Cardinality = Core.Cardinality.ecmSingleValued
        DefaultValue = lpDefaultValue
        SetPersistence(False)
        Description = lpDescription
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
