'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Defines a relationship between two or more documents.</summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Partial Public Class Relationship
    Implements IComparable



#Region "Public Methods"

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Just revert to the default behavior
        Return Me.GetType.Name
      End Try
    End Function

#End Region

#Region "Friend Methods"

    Public Function RemoveFromDocument(ByVal lpDocument As Document) As Boolean
      Try

        If lpDocument Is Nothing Then
          Throw New ArgumentNullException("lpDocument")
        End If

        If lpDocument.Relationships.Contains(Me) Then
          lpDocument.Relationships.Remove(Me)
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

    Protected Friend Function DebuggerIdentifier() As String
      Try

        Return String.Format("{0}. {1}:{2}-{3} -> {4}-{5}", Order, Type, ParentObjectType, ParentId, ChildObjectType, ChildId)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      If TypeOf obj Is Relationship Then
        Return ObjectID.CompareTo(obj.ObjectID)
      Else
        Throw New ArgumentException("Object is not a Relationship")
      End If

    End Function

#End Region

  End Class

End Namespace