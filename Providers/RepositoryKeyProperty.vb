'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  RepositoryKeyProperty.vb
'   Description :  [type_description_here]
'   Created     :  7/25/2014 6:31:48 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Providers

  ''' <summary>
  ''' A special type of provider property used to identify which repository(s) to establish a content source to.
  ''' </summary>
  ''' <remarks>
  ''' This is the only provider property which can contain multiple values.
  ''' <para></para>
  ''' <para>
  ''' If multiple values are assigned, the result should be that multiple content sources get created, one for each value.
  ''' </para>
  ''' </remarks>
  Public Class RepositoryKeyProperty
    Inherits ProviderProperty
    Implements IRepositoryKeyList

#Region "Overriden Properties"

    'Public Overrides Property Cardinality As Core.Cardinality
    '  Get
    '    Return Core.Cardinality.ecmMultiValued
    '  End Get
    '  Set(value As Core.Cardinality)
    '    ' Ignore set operations
    '  End Set
    'End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpPropertyName As String,
              Optional ByVal lpPropertyValue As String = "",
              Optional ByVal lpSequenceNumber As Integer = -1,
              Optional ByVal lpDescription As String = "")
      MyBase.New(lpPropertyName, GetType(String), True, lpPropertyValue, lpSequenceNumber, lpDescription, True, True)
    End Sub

#End Region

#Region "IRepositoryKeyList Implementation"

    Public Property SelectedKeys As Generic.List(Of String) Implements IRepositoryKeyList.SelectedKeys

#End Region

  End Class

End Namespace