'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>Action used to run another Transformation</summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class RunTransformationAction
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "RunTransformation"
    Friend Const PARAM_TRANSFORMATION_PATH As String = "TransformationPath"

#End Region

#Region "Class Variables"

    Private mstrTransformationPath As String
    Private mobjTransformation As Transformation

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    <XmlAttribute()>
    Public Overrides Property Name() As String
      Get
        Try

          'If String.IsNullOrEmpty(MyBase.Name) Then
          '  If TransformationPath IsNot Nothing AndAlso TransformationPath.Length > 0 Then
          '    mstrName = String.Format("{0}({1})", Me.GetType.Name, TransformationPath)
          '  Else
          '    mstrName = Me.GetType.Name
          '  End If
          'End If

          Return MyBase.Name

        Catch ex As Exception
          Return MyBase.mstrName
        End Try
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the path of the target transformation
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TransformationPath() As String
      Get
        Return mstrTransformationPath
      End Get
      Set(ByVal value As String)
        mstrTransformationPath = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the target transformation to be run
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Not serialized</remarks>
    <Xml.Serialization.XmlIgnore()>
    Public Property TargetTransformation() As Transformation
      Get
        Return mobjTransformation
      End Get
      Set(ByVal value As Transformation)
        mobjTransformation = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.RunTransformation)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(ActionType.RunTransformation, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpTransformationPath As String)
      MyBase.New(ActionType.RunTransformation, lpName)
      Try
        TransformationPath = lpTransformationPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpTargetTransformation As Transformation)
      MyBase.New(ActionType.RunTransformation, lpName)
      Try
        TransformationPath = String.Format("{0}{1}.ctf", Core.FileHelper.Instance.TempPath, lpName)
        TargetTransformation = lpTargetTransformation
        TargetTransformation.Serialize(TransformationPath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        Dim lobjActionResult As ActionResult = Nothing
        Dim lstrTransformationInitializationErrorMessage As String = String.Empty
        Dim lblnTransformationInitializationSuccess As Boolean

        If TargetTransformation Is Nothing Then
          lblnTransformationInitializationSuccess = InitializeTransformationFromPath(
            TransformationPath, lstrTransformationInitializationErrorMessage)
          If lblnTransformationInitializationSuccess = False Then
            Return New ActionResult(Me, False, lstrTransformationInitializationErrorMessage)
          End If
        End If

        If TargetTransformation IsNot Nothing Then
          If TypeOf Transformation.Target Is Document Then
            Transformation.Document = Transformation.Document.Transform(TargetTransformation)
            lobjActionResult = New ActionResult(Me, True)

            Return lobjActionResult
          ElseIf TypeOf Transformation.Target Is Folder Then
            Transformation.Folder = Transformation.Folder.Transform(TargetTransformation)
            lobjActionResult = New ActionResult(Me, True)

            Return lobjActionResult
          Else
            Throw New InvalidTransformationTargetException
          End If

        Else
          Return New ActionResult(Me, False,
                                  "Unable to execute RunTransformationAction, the TargetTransformation property is not set.")

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        If lpErrorMessage.Length = 0 Then
          lpErrorMessage = ex.Message
        End If
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try
    End Function

    Public Function FindTargetTransformation() As Transformation
      Try
        Return FindTargetTransformation(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FindTargetTransformation(ByVal lpErrorMessage As String) As Transformation
      Try
        'Dim lstrTransformationInitializationErrorMessage As String = String.Empty
        Dim lblnTransformationInitializationSuccess As Boolean
        If TargetTransformation IsNot Nothing Then
          Return TargetTransformation
        Else
          lblnTransformationInitializationSuccess = InitializeTransformationFromPath(
            TransformationPath, lpErrorMessage)
          If lblnTransformationInitializationSuccess = False Then
            Return Nothing
          Else
            Return TargetTransformation
          End If

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_TRANSFORMATION_PATH) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_TRANSFORMATION_PATH, String.Empty,
            "Specifies the path to the transformation to run."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        Me.TransformationPath = GetStringParameterValue(PARAM_TRANSFORMATION_PATH, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Friend Overloads Function DebuggerIdentifier() As String
      Try
        'If String.IsNullOrEmpty(Name) Then
        '  Return Me.GetType.Name
        'Else
        If String.Compare(Me.GetType.Name, Name) <> 0 Then
          Return Name
        Else
          Return String.Format("{0}({1})", Me.GetType.Name, TransformationPath)
          'String.Format("{0}({1})", Me.GetType.Name, TransformationPath)
        End If
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function InitializeTransformationFromPath(ByVal lpTransformationPath As String,
                                                      ByRef lpErrorMessage As String) As Boolean
      Try

        ' TODO: Add resolution to nearby paths such as \.. and FileHelper.TransformationPath 

        ' 1. See if we were given a fully qualified path.
        ' 2. See if it is in the same folder or a relative folder of the folder the original transformation was in.
        ' 3. See if it is in the 'Transformations' folder under the Cts Root path.

        Dim lstrPrimaryTransformationPath As String = Me.Transformation.TransformationFilePath
        Dim lstrTransformationPath As String = Nothing

        If IO.Path.IsPathRooted(lpTransformationPath) AndAlso IO.File.Exists(lpTransformationPath) Then
          ' This qualifies under scenario #1.
          lstrTransformationPath = lpTransformationPath
        End If

        If String.IsNullOrEmpty(lstrTransformationPath) Then
          If IO.Path.IsPathRooted(lpTransformationPath) = False AndAlso
             Not String.IsNullOrEmpty(Me.Transformation.TransformationFilePath) Then
            Dim lstrReferencedTransformationPath As String = String.Format("{0}\{1}",
                                IO.Path.GetDirectoryName(Me.Transformation.TransformationFilePath),
                                IO.Path.GetFileName(lpTransformationPath))
            If IO.File.Exists(lstrReferencedTransformationPath) Then
              ' This qualifies under scenario #2.
              lstrTransformationPath = lstrReferencedTransformationPath
            End If
          End If
        End If

        If String.IsNullOrEmpty(lstrTransformationPath) Then
          If IO.Path.IsPathRooted(lpTransformationPath) = False Then
            Dim lstrReferencedTransformationPath As String = String.Format("{0}\{1}",
                                FileHelper.Instance.TransformationPath,
                                IO.Path.GetFileName(lpTransformationPath))
            If IO.File.Exists(lstrReferencedTransformationPath) Then
              ' This qualifies under scenario #3.
              lstrTransformationPath = lstrReferencedTransformationPath
            End If
          End If
        End If

        If String.IsNullOrEmpty(lstrTransformationPath) Then
          ' We were not able to locate a matching file.
          lpErrorMessage = String.Format("Invalid TransformationPath, the file '{0}' could not be found.",
                          mstrTransformationPath)
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning)
          Return False
        Else
          ' We found a file
          Try
            TargetTransformation = New Transformation(lstrTransformationPath)
            Return True
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            lpErrorMessage = ex.Message
            Return False
          End Try
        End If

        '' Make sure the path is a valid file
        'If Not IO.File.Exists(lpTransformationPath) Then
        '  ' Try to load the transformation from the same 
        '  ' directory the parent transformation was loaded from.
        '  If String.IsNullOrEmpty(Me.Transformation.TransformationFilePath) Then
        '    lpErrorMessage = String.Format("Invalid TransformationPath, the file '{0}' could not be found.", _
        '                  mstrTransformationPath)
        '    ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning)
        '    Return False
        '  Else
        '    If IO.Path.IsPathRooted(lpTransformationPath) = False AndAlso _
        '      Not String.IsNullOrEmpty(Me.Transformation.TransformationFilePath) Then
        '      Dim lstrReferencedTransformationPath As String = String.Format("{0}\{1}", _
        '                      IO.Path.GetDirectoryName(Me.Transformation.TransformationFilePath), _
        '                      IO.Path.GetFileName(lpTransformationPath))
        '      If IO.File.Exists(lstrReferencedTransformationPath) Then
        '        Try
        '          TargetTransformation = New Transformation(lstrReferencedTransformationPath)
        '          Return True
        '        Catch ex As Exception
        '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '          lpErrorMessage = ex.Message
        '          Return False
        '        End Try
        '      End If
        '    End If
        '  End If
        '  lpErrorMessage = String.Format("Invalid TransformationPath, the file '{0}' could not be found.", _
        '                lpTransformationPath)
        '  ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning)
        '  Return False
        'End If

        '' Make sure the path is a valid transformation file
        'Try
        '  TargetTransformation = New Transformation(lpTransformationPath)
        '  Return True
        'Catch ex As Exception
        '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  lpErrorMessage = ex.Message
        '  Return False
        'End Try

        'Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = ex.Message
        Return False
      End Try
    End Function

#End Region

  End Class

End Namespace
