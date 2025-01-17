'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Providers

  ''' <summary>
  ''' Contains information about a provider
  ''' </summary>
  ''' <remarks>Intended for use in a provider catalog</remarks>
  Public Class ProviderInformation
    Implements IComparable

#Region "Class Variables"

    Private mstrId As String = String.Empty
    Private mstrSystemType As String = String.Empty
    Private mstrCompanyName As String = String.Empty
    Private mstrProductName As String = String.Empty
    Private mstrProductVersion As String = String.Empty
    Private mstrName As String = String.Empty
    Private mobjInterfaces As New List(Of String)
    Private mstrDescription As String = String.Empty
    Private mstrProviderPath As String = String.Empty
    Private mblnIsLicensed As Boolean
    Private mstrLicenseReason As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The Provider Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property
    ''' <summary>
    ''' The code name of the provider.  
    ''' This corresponds to the class 
    ''' name for the provider implementation.
    ''' </summary>
    <XmlAttribute()>
    Public ReadOnly Property Id As String
      Get
        Return mstrId
      End Get
    End Property

    <XmlAttribute()>
    Public Property IsLicensed() As Boolean
      Get
        Return mblnIsLicensed
      End Get
      Set(ByVal value As Boolean)
        If Helper.IsDeserializationBasedCall = True Then
          mblnIsLicensed = value
        Else
          Throw New InvalidOperationException(Core.TREAT_AS_READ_ONLY)
        End If
      End Set
    End Property

    <XmlAttribute()>
    Public Property LicenseDetail() As String
      Get
        Return mstrLicenseReason
      End Get
      Set(ByVal value As String)
        If Helper.IsDeserializationBasedCall = True Then
          mstrLicenseReason = value
        Else
          Throw New InvalidOperationException(Core.TREAT_AS_READ_ONLY)
        End If
      End Set
    End Property

    ''' <summary>
    ''' An optional property that describes the provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Description() As String
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

    ''' <summary>
    ''' The type of repository the provider connects to
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SystemType() As String
      Get
        Return mstrSystemType
      End Get
      Set(ByVal value As String)
        mstrSystemType = value
      End Set
    End Property

    ''' <summary>
    ''' The name of the company that authored the repository
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CompanyName() As String
      Get
        Return mstrCompanyName
      End Get
      Set(ByVal value As String)
        mstrCompanyName = value
      End Set
    End Property

    ''' <summary>
    ''' The product name of the repository the provider connects to
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProductName() As String
      Get
        Return mstrProductName
      End Get
      Set(ByVal value As String)
        mstrProductName = value
      End Set
    End Property

    ''' <summary>
    ''' The version of the repository that the provider targets
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This is often an optional value, 
    ''' but for some providers it is necessary 
    ''' due to version differences</remarks>
    Public Property ProductVersion() As String
      Get
        Return mstrProductVersion
      End Get
      Set(ByVal value As String)
        mstrProductVersion = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or Sets the Provider Path
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Set should only be called from a deserialization process</remarks>
    Public Property ProviderPath() As String
      Get
        Return mstrProviderPath
      End Get
      Set(ByVal value As String)
        Try
          If Helper.IsDeserializationBasedCall = True Then
            mstrProviderPath = value
          Else
            Throw New InvalidOperationException(Core.TREAT_AS_READ_ONLY)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The list of interfaces that this provider supports
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Interfaces() As List(Of String)
      Get
        Return mobjInterfaces
      End Get
      Set(ByVal value As List(Of String))
        mobjInterfaces = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a ProviderInformation object using the provider path
    ''' </summary>
    ''' <param name="lpProviderPath">A fully qualified provider path</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpProviderPath As String)
      Try
        InitializeFromProvider(ContentSource.GetProvider(lpProviderPath, Nothing, False))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Create a ProviderInformation object using the supplied Provider object
    ''' </summary>
    ''' <param name="lpProvider">A valid CProvider object reference</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpProvider As CProvider)
      Try
        InitializeFromProvider(lpProvider)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Public Methods"

    Public Shadows Function ToString() As String
      Try
        Return Name
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Sub InitializeFromProvider(ByVal lpProvider As CProvider)
      Try
        mstrId = lpProvider.GetType.Name
        Name = lpProvider.Name
        If lpProvider.ProviderSystem IsNot Nothing Then
          SystemType = lpProvider.ProviderSystem.Type
          CompanyName = lpProvider.ProviderSystem.Company
          ProductName = lpProvider.ProviderSystem.ProductName
          ProductVersion = lpProvider.ProviderSystem.ProductVersion
        End If
        mblnIsLicensed = lpProvider.HasValidLicense
        mstrLicenseReason = lpProvider.LicenseFailureReason
        GetInterfaces(lpProvider)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub GetInterfaces(ByVal lpProvider As CProvider)
      Try

        Dim lobjAssembly As System.Reflection.Assembly

        lobjAssembly = Reflection.Assembly.GetAssembly(lpProvider.GetType)

        mstrProviderPath = lobjAssembly.Location

        For Each lobjType As Type In lobjAssembly.GetTypes
          Debug.Print(lobjType.Name)
        Next

        For Each lobjType As System.Type In lobjAssembly.GetTypes

          If lobjType.Name.ToLower.EndsWith("provider") AndAlso lobjType.Name <> "ICustomAttributeProvider" Then
            For Each lobjInterface As System.Type In lobjType.GetInterfaces
              If lobjInterface.FullName.ToLower.StartsWith("ecmg.cts") Then
                Select Case lobjInterface.Name
                  Case "IProvider", "IAuthorization"
                    ' Skip these as we really don't care about them here
                  Case Else
                    mobjInterfaces.Add(lobjInterface.Name)
                End Select
              End If
            Next
            Exit Sub
          End If

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        If TypeOf obj Is ProviderInformation Then
          Return Name.CompareTo(obj.Name)
        Else
          Throw New ArgumentException("Object is not of type ProviderInformation")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace