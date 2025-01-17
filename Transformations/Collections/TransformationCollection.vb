'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class TransformationCollection
    Inherits CCollection(Of Transformation)
    Implements ICloneable

#Region "Class Variables"

    Private mobjPrimaryTransformation As Transformation = Nothing

#End Region

#Region "Public Properties"

    Public Property PrimaryTransformation As Transformation
      Get
        Return mobjPrimaryTransformation
      End Get
      Set(value As Transformation)
        mobjPrimaryTransformation = value
      End Set
    End Property

    Default Shadows Property Item(ByVal name As String) As Transformation
      Get
        Try
          For Each lobjTransformation As Transformation In Me
            If (lobjTransformation.Name.ToLower = name.ToLower) Then
              Return lobjTransformation
            End If
          Next

          Return Nothing

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As Transformation)
        Try
          Dim lobjProfile As Transformation
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProfile = CType(MyBase.Item(lintCounter), Transformation)
            If lobjProfile.Name = name Then
              MyBase.Item(lintCounter) = value
              Exit Property
            End If
          Next
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("TransformationCollection::Set_Item('{0}')", name))
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As Transformation
      Get
        Try
          Return MyBase.Item(Index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("TransformationCollection::Get_Item('{0}')", Index))
          ' Re-throw the exception to the caller
          Throw
        Finally
        End Try
      End Get
      Set(ByVal value As Transformation)
        MyBase.Item(Index) = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    Public Shadows Sub Add(lpTransformation As Transformation)
      Try
        If PrimaryTransformation Is Nothing Then
          PrimaryTransformation = lpTransformation
        End If
        MyBase.Add(lpTransformation)
        ' Find any child transformations if they exist and add them as well.
        MyBase.Add(lpTransformation.GetChildTransformations)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone
      Try
        'Return New Process(Me)
        Dim lstrProcessString As String = Serializer.Serialize.XmlString(Me)
        Return Serializer.Deserialize.XmlString(lstrProcessString, GetType(TransformationCollection))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
