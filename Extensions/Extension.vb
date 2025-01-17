' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Extension.vb
'  Description :  Used for managing extension dlls for projects.
'  Created     :  11/16/2011 7:34:46 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents.SerializationUtilities
Imports Documents.Utilities


#End Region

Namespace Extensions

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Extension
    Implements IExtension
    Implements IDisposable

#Region "Class Constants"

    Private Const IEXTENSION_NAME As String = "IExtension"

    Protected NAME_CONST As String = "Name"
    Protected DESC_CONST As String = "Description"
    Protected COMPANY_NAME_CONST As String = "Company Name"
    Protected PRODUCT_NAME_CONST As String = "Product Name"

#End Region

#Region "Class Variables"

    Private mstrAddedByMachine As String = String.Empty
    Private mstrAddedByUser As String = String.Empty
    Private mobjByteArray As Byte()
    Private mdatCreateDate As Date = Date.Now
    Private mstrId As String = String.Empty
    Private mstrPath As String = String.Empty
    Private mstrName As String = NAME_CONST
    Private mstrDisplayName As String = NAME_CONST
    Private mstrDescription As String = NAME_CONST
    Private mstrCompanyName As String = String.Empty
    Private mstrProductName As String = String.Empty
    Private mblnIsValid As Boolean = True
    Private mobjLoadException As Exception = Nothing

#End Region

#Region "IExtension Implementation"

    ''' <summary>
    ''' Gets or sets the machine name that the extension was originally added by.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AddedByMachine As String Implements IExtension.AddedByMachine
      Get
        Return mstrAddedByMachine
      End Get
      Set(value As String)
        mstrAddedByMachine = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the user that added the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AddedByUser As String Implements IExtension.AddedByUser
      Get
        Return mstrAddedByUser
      End Get
      Set(value As String)
        mstrAddedByUser = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a byte array containing the extension assembly.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ByteArray As Byte() Implements IExtension.ByteArray
      Get
        Return mobjByteArray
      End Get
      Set(value As Byte())
        mobjByteArray = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the date the extension was added.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CreateDate As Date Implements IExtension.CreateDate
      Get
        Return mdatCreateDate
      End Get
      Set(value As Date)
        mdatCreateDate = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the identifier for the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Id As String Implements IExtension.Id
      Get
        Return mstrId
      End Get
      Set(value As String)
        mstrId = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the name of the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property Name As String Implements IExtension.Name
      Get
        Return mstrName
      End Get
      'Set(value As String)
      '  mstrName = value
      'End Set
    End Property

    Public Overridable ReadOnly Property DisplayName As String Implements IExtension.DisplayName
      Get
        Try
          If String.Equals(NAME_CONST, mstrDisplayName) Then
            If String.Equals(NAME_CONST, mstrName) Then
              Return mstrDisplayName
            Else
              Return Helper.CreateDisplayName(Name)
            End If
          Else
            Return Helper.CreateDisplayName(Name)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the description of the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property Description As String Implements IExtension.Description
      Get
        Return mstrDescription
      End Get
      'Set(value As String)
      '  mstrDescription = value
      'End Set
    End Property

    Public Overridable ReadOnly Property CompanyName As String Implements IExtension.CompanyName
      Get
        Try
          If String.IsNullOrEmpty(mstrCompanyName) Then
            If Assembly.GetCallingAssembly.FullName.ToUpper.Contains("ECMG") Then
              mstrCompanyName = Core.ConstantValues.CompanyName
            Else
              mstrCompanyName = COMPANY_NAME_CONST
            End If
          End If
          Return mstrCompanyName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overridable ReadOnly Property ProductName As String Implements IExtension.ProductName
      Get
        Try
          If String.IsNullOrEmpty(mstrProductName) Then
            If Assembly.GetCallingAssembly.FullName.ToUpper.Contains("ECMG") Then
              mstrProductName = Core.ConstantValues.ProductName
            Else
              mstrProductName = PRODUCT_NAME_CONST
            End If
          End If
          Return mstrProductName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overridable ReadOnly Property Path As String Implements IExtension.Path
      Get
        Return mstrPath
      End Get
    End Property

    Public ReadOnly Property IsValid As Boolean Implements IExtensionInformation.IsValid
      Get
        Try
          Return mblnIsValid
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property LoadException As Exception Implements IExtensionInformation.LoadException
      Get
        Try
          Return mobjLoadException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend Sub OnLoadError(exception As Exception) Implements IExtensionInformation.OnLoadError
      Try
        If exception Is Nothing Then
          Throw New ArgumentNullException("exception")
        End If
        mobjLoadException = exception
        mblnIsValid = False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlElementString() As String Implements IExtensionInformation.ToXmlElementString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    ''' <summary>
    ''' The default constructor is not public.  All construction of 
    ''' extensions should go throw the CreateExtension method.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()

    End Sub

    Protected Sub New(lpPath As String)
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Protected Sub new(lpName As String, lpDescription As String, 
#End Region

#Region "Public Shared Methods"

    Public Shared Function CreateExtension(ByVal lpName As String,
                                       ByVal lpPath As String) As IExtension
      Try
        Dim lobjAssembly As Assembly = Assembly.LoadFrom(lpPath)

        Return CreateExtension(lpName, lobjAssembly)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateExtension(ByVal lpName As String,
                                       ByVal lpAssembly As Assembly) As IExtension
      Try
        'Return CreateExtension(String.Empty, lpName, lpDescription, _
        '                       AssemblyHelper.WriteAssemblyToByteArray(lpAssembly))

        Dim lobjTypes As List(Of Type) = Nothing
        Dim lobjInstance As IExtension = Nothing

        ' Get all the extension types from the assembly
        lobjTypes = Helper.GetTypesFromAssembly(lpAssembly, GetType(IExtension))

        For Each lobjType As Type In lobjTypes
          lobjInstance = Activator.CreateInstance(lobjType)

          If String.Compare(lpName, lobjInstance.Name, True) Then
            Return lobjInstance
          End If
        Next

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Factory method for creating an extension object.
    ''' </summary>
    ''' <param name="lpName">The name of the extension.</param>
    ''' <param name="lpDescription">The description of the extension.</param>
    ''' <param name="lpAssembly">The assembly as a byte array.</param>
    ''' <returns>A new IExtension object reference.</returns>
    ''' <remarks></remarks>
    Public Shared Function CreateExtension(ByVal lpName As String,
                                       ByVal lpDescription As String,
                                       ByVal lpAssembly() As Byte) As IExtension
      Try
        Return CreateExtension(String.Empty, lpName, lpDescription, lpAssembly)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Factory method for creating an extension object.
    ''' </summary>
    ''' <param name="lpId">The id of the extension.</param>
    ''' <param name="lpName">The name of the extension.</param>
    ''' <param name="lpDescription">The description of the extension.</param>
    ''' <param name="lpAssembly">The assembly as a byte array.</param>
    ''' <returns>A new IExtension object reference.</returns>
    ''' <remarks></remarks>
    Public Shared Function CreateExtension(ByVal lpId As String,
                                           ByVal lpName As String,
                                           ByVal lpDescription As String,
                                           ByVal lpAssembly() As Byte) As IExtension
      Try
        Dim lobjExtension As New Extension

        With lobjExtension
          .Id = lpId
          '.Name = lpName
          'If Not String.IsNullOrEmpty(lpDescription) Then
          '  .Description = lpDescription
          'End If
          .ByteArray = lpAssembly
          .AddedByMachine = Environment.MachineName
          .AddedByUser = Environment.UserName
          .CreateDate = Now
        End With

        Return lobjExtension

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New Text.StringBuilder
      Try

        If disposedValue = True Then
          If String.IsNullOrEmpty(Me.Name) Then
            Return "Extension Disposed"
          Else
            Return String.Format("{0} Extension Disposed", Me.Name)
          End If
        End If

        lobjIdentifierBuilder.AppendFormat("{0} Extension", Me.GetType.Name)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Sub SetName(lpName As String)
      Try
        mstrName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetDisplayName(lpDisplayName As String)
      Try
        mstrDisplayName = lpDisplayName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetDescription(lpDescription As String)
      Try
        mstrDescription = lpDescription
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

#End Region

#Region " IDisposable Support "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: free other state (managed objects).
            mstrAddedByMachine = Nothing
            mstrAddedByUser = Nothing
            mobjByteArray = Nothing
            mdatCreateDate = Nothing
            mstrId = Nothing
            mstrPath = Nothing
            mstrName = Nothing
            mstrDisplayName = Nothing
            mstrDescription = Nothing
          End If

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' TODO: set large fields to null.

        End If
        Me.disposedValue = True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

  End Class

End Namespace