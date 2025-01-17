'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ExportObjectEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  9/1/2015 3:33:20 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class ExportObjectEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mstrObjectId As String = String.Empty
    Private mobjObject As CustomObject = Nothing
    Private mblnGetPermissions As Boolean = True





#End Region
#Region "Public Properties"

    Public Property [Object] As CustomObject
      Get
        Try
          Return mobjObject
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As CustomObject)
        Try
          mobjObject = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property ObjectId() As String
      Get
        Return mstrObjectId
      End Get
    End Property


    ''' <summary>Determines whether or not the export operation should get or exclude the document permissions</summary>
    ''' <remarks>This property can only be used on exports with providers that implement ISecurityClassification.</remarks>
    ''' <seealso cref="Providers.ISecurityClassification">ISecurityClassification</seealso>
    Public Property GetPermissions As Boolean
      Get
        Return mblnGetPermissions
      End Get
      Set(value As Boolean)
        mblnGetPermissions = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      MyBase.New()
      Try
        mstrObjectId = lpId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace