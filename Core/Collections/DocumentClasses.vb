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

  ''' <summary>Collection of DocumentClass objects.</summary>
  Public Class DocumentClasses
    Inherits RepositoryObjectClasses
    'Inherits CCollection(Of DocumentClass)
    'Implements INamedItems

#Region "Constructors"

    Public Sub New()
      MyBase.new()
    End Sub

    Public Sub New(lpObjectClasses As RepositoryObjectClasses)
      Try

        For Each lobjObjectClass As RepositoryObjectClass In lpObjectClasses
          If TypeOf lobjObjectClass Is DocumentClass Then
            MyBase.Add(lobjObjectClass)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function CoalesceAllProperties() As ClassificationProperties
      Try
        Dim lobjAllProperties As New ClassificationProperties

        For Each lobjDocumentClass As DocumentClass In Me
          For Each lobjProperty As ClassificationProperty In lobjDocumentClass.Properties
            If Not lobjAllProperties.Contains(lobjProperty.Name) Then
              lobjAllProperties.Add(lobjProperty)
            End If
          Next
        Next

        lobjAllProperties.SortCollection("Name", True)

        Return lobjAllProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub Add(ByVal item As DocumentClass)
      Try
        MyBase.Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal items As DocumentClasses)
      Try
        MyBase.Add(items)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Function ClassExists(ByVal lpClassNameOrId As String) As Boolean
    '  '  Does the DocumentClass exist?
    '  Try
    '    'Dim lobjDocumentClass As DocumentClass = Item(lpClassName)
    '    'Return True

    '    For Each lobjDocumentClass As DocumentClass In Me
    '      If StrComp(lobjDocumentClass.Name, lpClassNameOrId, CompareMethod.Binary) = 0 Then
    '        Return True
    '      End If
    '    Next

    '    For Each lobjDocumentClass As DocumentClass In Me
    '      If StrComp(lobjDocumentClass.ID, lpClassNameOrId, CompareMethod.Binary) = 0 Then
    '        Return True
    '      End If
    '    Next

    '    Return False

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Return False
    '  End Try

    'End Function

    'Public Overloads Function Contains(ByVal lpClassName As String) As Boolean

    '  ' First look for the class by Name
    '  For Each lobjDocumentClass As DocumentClass In Me
    '    If lobjDocumentClass.Name = lpClassName Then
    '      Return True
    '    End If
    '  Next

    '  ' Now look for the class by Label
    '  For Each lobjDocumentClass As DocumentClass In Me
    '    If lobjDocumentClass.Label = lpClassName Then
    '      Return True
    '    End If
    '  Next

    '  Return False

    'End Function

    ''' <summary>
    ''' Gets or sets the document class based on the specified identifier.
    ''' </summary>
    ''' <param name="Name">Interpreted in order as Name, Label and ID.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Shadows Property Item(ByVal Name As String) As DocumentClass
      Get
        Dim lobjDocumentClass As DocumentClass
        ' First look for the class by Name
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjDocumentClass
          End If
        Next

        ' Next look for the class by Label
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.Label, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjDocumentClass
          End If
        Next

        ' Next look for the class by ID
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.ID, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjDocumentClass
          End If
        Next

        Return Nothing

      End Get
      Set(ByVal value As DocumentClass)
        Dim lobjDocumentClass As DocumentClass

        ' First look for the class by Name
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
            MyBase.Item(lintCounter) = value
            Exit Property
          End If
        Next

        ' Next look for the class by Label
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.Label, Name, StringComparison.InvariantCultureIgnoreCase) Then
            MyBase.Item(lintCounter) = value
            Exit Property
          End If
        Next

        ' Next look for the class by ID
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjDocumentClass = CType(MyBase.Item(lintCounter), DocumentClass)
          If String.Equals(lobjDocumentClass.ID, Name, StringComparison.InvariantCultureIgnoreCase) Then
            MyBase.Item(lintCounter) = value
            Exit Property
          End If
        Next

        Throw New Exception("There is no Item by the Name '" & Name & "'.")
      End Set
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal name As String,
                                                  ByVal ignoreCase As Boolean) As DocumentClass
      Get
        Return MyBase.Item(name, ignoreCase)
      End Get
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As DocumentClass
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As DocumentClass)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Shadows Function Remove(ByVal Item As DocumentClass) As Boolean
      Try
        ' Since sometimes we could have two different object references 
        ' to the same class, we will attempt to handle it here by the 
        ' assumption that if the class has the same name, then it is 
        ' the one we really care about.
        If Contains(Item.Name) Then
          Return MyBase.Remove(Me.Item(Item.Name))
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Function ToXmlString() As String

    '  Return Serializer.Serialize.XmlString(Me)

    'End Function

    Protected Overrides Sub Finalize()
      MyBase.Finalize()
    End Sub

#End Region

    '#Region "INamedItems Implementation"

    '    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
    '      Try
    '        If TypeOf (item) Is DocumentClass Then
    '          Add(DirectCast(item, DocumentClass))
    '        Else
    '          Throw New ArgumentException( _
    '            String.Format("Invalid object type.  Item type of DocumentClass expected instead of type {0}", _
    '                          item.GetType.Name), "item")
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Property ItemByName(ByVal name As String) As Object Implements INamedItems.ItemByName
    '      Get
    '        Return Item(name)
    '      End Get
    '      Set(ByVal value As Object)
    '        Item(name) = value
    '      End Set
    '    End Property

    '    Public Shadows Function Remove(ByVal lpName As String) As Boolean Implements INamedItems.Remove
    '      Try
    '        ' Since sometimes we could have two different object references 
    '        ' to the same class, we will attempt to handle it here by the 
    '        ' assumption that if the class has the same name, then it is 
    '        ' the one we really care about.
    '        If Contains(lpName) Then
    '          Return MyBase.Remove(Me.Item(lpName))
    '        Else
    '          Return False
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Function

    '#End Region

  End Class

End Namespace