'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MapListParameter.vb
'   Description :  [type_description_here]
'   Created     :  5/5/2014 8:30:44 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class MapListParameter
    Inherits MultiValueParameter

#Region "Class Variables"

    Private WithEvents mobjMapList As New MapList

#End Region

#Region "Public Properties"

    Public Property MapList As MapList
      Get
        Try
          Return mobjMapList
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As MapList)
        Try
          mobjMapList = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmValueMap)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmValueMap, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(PropertyType.ecmValueMap, lpName, lpSystemName, lpValues)
      Try
        If Values IsNot Nothing Then
          If Values.Count > 0 Then
            If TypeOf Values.First Is ValueMap Then
              For Each lobjValue As ValueMap In Values
                mobjMapList.Add(lobjValue)
              Next
            End If
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Sub mobjMapList_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles mobjMapList.CollectionChanged
      Try
        SetValuesFromMapList()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub SetValuesFromMapList()
      Try
        Values.Clear()
        If mobjMapList IsNot Nothing Then
          Values.AddRange(mobjMapList)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace