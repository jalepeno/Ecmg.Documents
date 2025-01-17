'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Base class from which many of the Cts object types inherit.</summary>
  Public MustInherit Class CtsObject
    Implements IDescription

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrConnectionString As String = String.Empty
    Private mstrUserName As String = String.Empty
    Private mstrPassword As String = String.Empty

#End Region

#Region "Public Properties"

    Public Property Name() As String Implements IDescription.Name, INamedItem.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Property Description() As String Implements IDescription.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

    ''' <summary>The connection string used to create the object.</summary>
    ''' <value>A string value representing the connection string.</value>
    Public Overridable Property ConnectionString() As String
      Get
        Return mstrConnectionString
      End Get
      Set(ByVal value As String)
        mstrConnectionString = value
      End Set
    End Property

    ''' <summary>The UserName associated with the current object.</summary>
    Public Overridable Property UserName() As String
      Get
        Return mstrUserName
      End Get
      Set(ByVal Value As String)
        mstrUserName = Value
      End Set
    End Property

    ' <XmlIgnore()> _
    ''' <summary>The password associated with the current object.</summary>
    Public Overridable Property Password() As String
      Get
        Return mstrPassword
      End Get
      Set(ByVal Value As String)
        mstrPassword = Value
      End Set
    End Property

    ' <XmlElement("Password")> _
    'Public Property EncryptedPassword() As String
    '  Get
    '    Return Uri.EscapeDataString(Core.Password.Encrypt(mstrPassword))
    '  End Get
    '  Set(ByVal Value As String)
    '    mstrPassword = DecryptPassword(Value)
    '  End Set
    'End Property

    'Public ReadOnly Property DecryptedPassword() As String
    '  Get
    '    Return Uri.UnescapeDataString(Core.Password.Decrypt(Password))
    '  End Get
    'End Property

    'Private Function DecryptPassword(ByVal lpPassword As String) As String
    '  Return Core.Password.Decrypt(Uri.UnescapeDataString(lpPassword))
    'End Function

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpConnectionString As String)

      Try
        If lpConnectionString.Length > 0 Then
          mstrConnectionString = lpConnectionString
          ParseConnectionString()
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CtsObject::New('{1}')", Me.GetType.Name, lpConnectionString))
      End Try

    End Sub

#End Region

#Region "Private methods"

    Private Sub ParseConnectionString()

      Try
        ' Get the UserName
        mstrUserName = Helper.GetInfoFromString(mstrConnectionString, "UserName")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CtsObject::ParseConnectionString_GetUserName('{0}')", mstrConnectionString))
      End Try

      If UserName.Length = 0 Then
        ' Try it this way
        Try
          ' Get the UserName
          mstrUserName = Helper.GetInfoFromString(mstrConnectionString, "uid")

        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("CtsObject::ParseConnectionString_GetUID('{0}')", mstrConnectionString))
        End Try
      End If


      Try
        ' Get the Password
        Password = Helper.GetInfoFromString(mstrConnectionString, "Password")
        Try
          Dim lobjUnEncryptedPassword As String = Core.Password.EscapedDecrypt(Password)
          Password = lobjUnEncryptedPassword
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("CtsObject::ParseConnectionString_GetEscapedDecryptedPassword('{0}')", mstrConnectionString))
          ' Leave it as it was
        End Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CtsObject::ParseConnectionString_GetPassword('{0}')", mstrConnectionString))
      End Try

      If Password.Length = 0 Then
        ' Try it this way
        Try
          ' Get the Password
          Password = Helper.GetInfoFromString(mstrConnectionString, "pwd")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("CtsObject::ParseConnectionString_GetPwd('{0}')", mstrConnectionString))
        End Try
      End If
    End Sub

#End Region

  End Class

End Namespace