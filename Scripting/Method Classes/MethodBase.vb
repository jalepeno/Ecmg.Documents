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

Namespace Scripting

  Public Class MethodBase
    Implements IMetaHolder
    Implements IMethod

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mobjParameters As New ECMProperties
    'Protected mobjLastResult As New MethodResult

#End Region

#Region "Public Properties"

    Public Property Name() As String Implements IMethod.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Property Parameters() As ECMProperties Implements IMethod.Parameters
      Get
        Return mobjParameters
      End Get
      Set(ByVal value As ECMProperties)
        mobjParameters = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub


#Region "Public Methods"

    'Public MustOverride Function Execute() As MethodResult Implements IMethod.Execute

    'Public MustOverride Function LastResult() As MethodResult Implements IMethod.LastResult

#End Region

#End Region

#Region "IMetaHolder Implementation"

    Public Property Identifier() As String Implements Core.IMetaHolder.Identifier
      Get
        Return Name
      End Get
      Set(ByVal value As String)
        Name = value
      End Set
    End Property

    Public Property Metadata() As Core.IProperties Implements Core.IMetaHolder.Metadata
      Get
        Return Parameters
      End Get
      Set(ByVal value As Core.IProperties)
        Parameters = value
      End Set
    End Property

#End Region

  End Class

  ''' <summary>
  ''' A collection of method objects inherited from MethodBase
  ''' </summary>
  ''' <remarks></remarks>
  Public Class Methods
    Inherits Core.CCollection(Of MethodBase)

  End Class

  Public Class MethodResult
    Inherits Core.ActionResult

#Region "Class Variables"

    Private mobjMethod As MethodBase
    Private mobjReturnValue As String = String.Empty

#End Region

#Region "Public Properties"

    Public Property Method() As MethodBase
      Get
        Return mobjMethod
      End Get
      Set(ByVal value As MethodBase)
        mobjMethod = value
      End Set
    End Property

    Public ReadOnly Property ReturnValue() As String
      Get
        Return mobjReturnValue
      End Get
    End Property
#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpMethod As MethodBase,
                   ByVal lpSuccess As Boolean,
                   Optional ByVal lpDetails As String = "")
      MyBase.New(lpSuccess, lpDetails)
      Method = lpMethod
    End Sub

    Public Sub New(ByVal lpMethod As MethodBase,
                   ByVal lpSuccess As Boolean,
                   ByVal lpReturnValue As Object,
                   Optional ByVal lpDetails As String = "")
      MyBase.New(lpSuccess, lpDetails)
      Method = lpMethod
      mobjReturnValue = lpReturnValue
    End Sub

#End Region

  End Class


  Public Class Parameters
    Inherits List(Of KeyValuePair(Of String, Object))

    Public Overloads Sub Add(ByVal lpParameterName As String, ByVal lpParameterValue As Object)
      MyBase.Add(New KeyValuePair(Of String, Object)(lpParameterName, lpParameterValue))
    End Sub

    Default Public Overloads ReadOnly Property Item(ByVal lpParameterName As String) As KeyValuePair(Of String, Object)
      Get
        Try
          For Each lobjKeyValuePair As KeyValuePair(Of String, Object) In Me
            If lobjKeyValuePair.Key = lpParameterName Then
              Return lobjKeyValuePair
            End If
          Next

          Return Nothing

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

      End Get

    End Property

  End Class

End Namespace