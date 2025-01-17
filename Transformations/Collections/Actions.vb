'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.Utilities

Namespace Transformations
  ''' <summary>Collection of Action objects.</summary>
  <Serializable()>
  Public Class Actions
    Inherits CCollection(Of Action)
    Implements ITransformationActions

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of ITransformationAction)

#End Region

#Region "Public Properties"

    Public Property Transformation As Transformation

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpTransformation As Transformation)
      MyBase.New()
      Try
        Transformation = lpTransformation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

#Region "Add Overrides"

    Public Shadows Sub Add(ByVal lpAction As Action)
      Try
        If lpAction Is Nothing Then
          ApplicationLogging.WriteLogEntry("Unable to add action, a null value was passed", TraceEventType.Warning)
          Exit Sub
        End If
        lpAction.Transformation = Me.Transformation
        MyBase.Add(lpAction)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>Adds the specified set of transformation actions to the collection.</summary>
    Public Shadows Sub Add(ByVal lpActions As Actions)

      Try
        For Each lobjAction As Action In lpActions
          MyBase.Add(lobjAction)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

    ''' <summary>
    ''' Looks through all actions and if a RunTransformationAction 
    ''' is present, replaces it with all the actions in the target 
    ''' transformation.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Combine()
      Try

        Dim lobjCombinedActions As New Actions

        For Each lobjAction As Action In Me
          If TypeOf lobjAction Is RunTransformationAction Then
            Dim lobjRunActionTarget As Transformation = DirectCast(lobjAction, RunTransformationAction).TargetTransformation
            If lobjRunActionTarget IsNot Nothing Then
              lobjCombinedActions.Add(lobjRunActionTarget.Combine.Actions)
            End If
          Else
            lobjCombinedActions.Add(lobjAction)
          End If
        Next

        ClearItems()

        Add(lobjCombinedActions)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function GetChildTransformations() As IList(Of Transformation)
      Try

        Dim lobjChildTransformations As New List(Of Transformation)
        Dim lobjRunTransformationAction As RunTransformationAction = Nothing
        Dim lobjRunActionTarget As Transformation = Nothing

        For Each lobjAction As Action In Me
          If TypeOf lobjAction Is RunTransformationAction Then
            lobjRunActionTarget = Nothing
            lobjRunTransformationAction = lobjAction
            If lobjRunTransformationAction.TargetTransformation Is Nothing Then
              lobjRunTransformationAction.TargetTransformation = lobjRunTransformationAction.FindTargetTransformation
            End If
            lobjRunTransformationAction.TargetTransformation = lobjRunActionTarget
            If lobjRunActionTarget IsNot Nothing Then
              lobjChildTransformations.Add(lobjRunActionTarget)
            End If
          End If
        Next

        Return lobjChildTransformations

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Public Shadows Sub Add(item As ITransformationAction) Implements ICollection(Of ITransformationAction).Add
      Try
        If TypeOf item Is Action Then
          Add(CType(item, Action))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Clear() Implements ICollection(Of ITransformationAction).Clear
      Try
        MyBase.Clear
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function Contains(item As ITransformationAction) As Boolean Implements ICollection(Of ITransformationAction).Contains
      Try
        If TypeOf item Is Action Then
          Return MyBase.Contains(CType(item, Action))
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub CopyTo(array() As ITransformationAction, arrayIndex As Integer) Implements ICollection(Of ITransformationAction).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows ReadOnly Property Count As Integer Implements ICollection(Of ITransformationAction).Count
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shadows ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of ITransformationAction).IsReadOnly
      Get
        Try
          Return MyBase.IsReadOnly
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shadows Function Remove(item As ITransformationAction) As Boolean Implements ICollection(Of ITransformationAction).Remove
      Try
        If TypeOf item Is Action Then
          Return MyBase.Remove(CType(item, Action))
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function GetEnumerator() As IEnumerator(Of ITransformationAction) Implements IEnumerable(Of ITransformationAction).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ItemById(lpId As String) As Action
      Try
        For Each lobjItem As Action In Me.Items
          If lobjItem.Id.Equals(lpId) Then
            Return lobjItem
          End If
        Next

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ToActionItems() As IActionItems
      Try
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of ITransformationAction)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of ITransformationAction)(Me.ToArray, GetType(Action), GetType(ITransformationAction))
          End If
          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

  End Class

End Namespace