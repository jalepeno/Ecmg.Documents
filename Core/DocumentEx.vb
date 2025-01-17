'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Security.Principal
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Arguments
Imports Documents.Exceptions
Imports Documents.Migrations
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Transformations
Imports Documents.Utilities
Imports Ionic.Zip
Imports Newtonsoft.Json


#End Region

Namespace Core

  <Serializable()>
  Partial Public Class Document
    Implements ISerialize
    Implements ICloneable
    Implements IDisposable
    Implements IMetaHolder
    Implements ITableable
    'Implements ILoggable

#Region "Class Enumerations"

    ''' <summary>Specifies which Hash is to be verified</summary>
    Public Enum HashLevel
      ''' <summary>Original Hash</summary>
      Original = 0
      ''' <summary>Latest Hash</summary>
      Last = 1
    End Enum

    ''' <summary>
    ''' Specifies how to treat ContentPaths for each child Content object
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ContentReferenceBehavior
      ''' <summary>Leaves the ContentPath values unchanged</summary>
      DoNotChange = 0
      ''' <summary>Moves the Content under version specific subdirectory of the Document Serialization Path</summary>
      Move = 1
      ''' <summary>Copies the Content under version specific subdirectory of the Document Serialization Path</summary>
      Copy = 2
    End Enum

#End Region

#Region "Member Classes"

    ''' <summary>
    ''' Encapsulates all of the document header information
    ''' </summary>
    ''' <remarks></remarks>
    ''' <exclude/>
    <Microsoft.VisualBasic.HideModuleName()>
    <Serializable()>
    <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
    Public Class DocumentHeader

#Region "Class Constants"

      Friend Const NO_ORIGINAL_HASH As String = "No Original Hash"

#End Region

#Region "Class Variables"

      Private mstrID As String = String.Empty
      Private mstrCtsVersion As String = String.Empty
      Private mstrOriginalLocale As String = Globalization.CultureInfo.CurrentCulture.Name 'My.Computer.Info.InstalledUICulture.Name
      Private mstrSerializationPath As String = String.Empty
      '<NonSerializedAttribute()>
      'Private mobjContentSource As ContentSource
      'Private mstrContentSourceConnectionString As String = String.Empty
      Private mdatGenerationDate As DateTime
      Private mdatGenenerationUtcDate As DateTime
      Private mstrOriginalHash As String = String.Empty
      Private mstrWorkstation As String = String.Empty
      Private mstrUserID As String = String.Empty
      Private mobjTransformationSeries As New Modifications
      Private mobjProperties As New HeaderProperties

#End Region

#Region "Public Properties"

      ''' <summary>
      ''' The document ID
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property ID() As String
        Get
          Return mstrID
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrID = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The version of Cts that was originally used to create or export the document
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property CtsVersion() As String
        Get
          Return mstrCtsVersion
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrCtsVersion = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The name of the locale used in the 
      ''' original creation or export of the document.
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      <XmlAttribute("OriginalLocale")>
      Public Property OriginalLocale() As String
        Get
          Try
            Return mstrOriginalLocale
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            Return String.Empty
          End Try
        End Get
        Set(ByVal value As String)
          mstrOriginalLocale = value
        End Set
      End Property

      Public Property SerializationPath() As String
        Get
          Return mstrSerializationPath
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrSerializationPath = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The date and time the document was exported
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property GenerationDate() As DateTime
        Get
          Return mdatGenerationDate
        End Get
        Set(ByVal value As DateTime)
          Try
            If Helper.IsDeserializationBasedCall Then
              mdatGenerationDate = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      Public Property GenenerationUtcDate() As DateTime
        Get
          Return mdatGenenerationUtcDate
        End Get
        Set(ByVal value As DateTime)
          Try
            If Helper.IsDeserializationBasedCall Then
              mdatGenenerationUtcDate = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      Public Property OriginalHash() As String
        Get
          Return mstrOriginalHash
        End Get
        Set(ByVal value As String)
          Try
            If mstrOriginalHash.Length = 0 Then
              mstrOriginalHash = value
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property
      ''' <summary>
      ''' The name of the computer the cdf file was generated with
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks>This is not the name of the repository server, but rather the name of the workstation that generated the document export</remarks>
      Public Property Workstation() As String
        Get
          Return mstrWorkstation
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrWorkstation = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The network user id of the person that generated the cdf file
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property UserID() As String
        Get
          Try
            If mstrUserID.Length = 0 Then
              mstrUserID = WindowsIdentity.GetCurrent.Name 'FileHelper.CurrentUser
            End If
            Return mstrUserID
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrUserID = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      '''' <summary>
      '''' The ContentSource used to export the document
      '''' </summary>
      '''' <value></value>
      '''' <returns></returns>
      '''' <remarks>If the document was created as a new document in the absence of a repository, then this property should be empty.</remarks>
      'Public Property ContentSource() As ContentSource
      '  Get
      '    Return mobjContentSource
      '  End Get
      '  Set(ByVal value As ContentSource)
      '    Try
      '      If Helper.IsDeserializationBasedCall Then
      '        mobjContentSource = value
      '      Else
      '        Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
      '      End If
      '    Catch ex As Exception
      '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '      ' Re-throw the exception to the caller
      '      Throw
      '    End Try
      '  End Set
      'End Property

      'Public Property ContentSourceConnectionString() As String
      '  Get
      '    Try
      '      If mstrContentSourceConnectionString.Length = 0 AndAlso
      '        mobjContentSource IsNot Nothing Then
      '        mstrContentSourceConnectionString = mobjContentSource.ConnectionString
      '      End If

      '      Return mstrContentSourceConnectionString
      '    Catch ex As Exception
      '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '      ' Re-throw the exception to the caller
      '      Throw
      '    End Try
      '  End Get
      '  Set(ByVal value As String)
      '    Try
      '      If Helper.IsDeserializationBasedCall Then
      '        mstrContentSourceConnectionString = value
      '      Else
      '        Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
      '      End If
      '    Catch ex As Exception
      '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '      ' Re-throw the exception to the caller
      '      Throw
      '    End Try
      '  End Set
      'End Property

      Public Property Properties() As HeaderProperties
        Get
          Return mobjProperties
        End Get
        Set(ByVal value As HeaderProperties)
          Try
            If Helper.IsDeserializationBasedCall Then
              mobjProperties = value
            Else
              Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The history of all of the transformations this document has been through
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property TransformationSeries() As Modifications
        Get
          Return mobjTransformationSeries
        End Get
        Set(ByVal value As Modifications)
          Try
            If Helper.IsDeserializationBasedCall Then
              mobjTransformationSeries = value
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

#End Region

      Friend Function GenerateHeaderString() As String

        Dim lstrHeaderString As String
        Dim lstrCompressedXML As String

        Try

          lstrHeaderString = Me.ToString
          lstrCompressedXML = Compression.StringCompression.CompressString(lstrHeaderString)
          Return Password.Encrypt(lstrCompressedXML).Hex

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' We failed to generate a header string
          ApplicationLogging.WriteLogEntry(
            String.Format(
              "Failed to generate header string. The serialization of the document header failed with the following error: {0}",
              ex.Message),
            TraceEventType.Error, 8975)
          Return String.Empty
        End Try

      End Function

      Friend Shared Function GenerateHeader(ByVal lpHeaderString As String) As DocumentHeader
        Try

          Dim lstrCompressedXML As String = Core.Password.DecryptFromHex(lpHeaderString).Text
          Dim lstrUnCompressedXML As String = Compression.StringCompression.DecompressString(lstrCompressedXML)
          Dim lobjHeader As New DocumentHeader(lstrUnCompressedXML)

          Return lobjHeader

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

      Friend Function Deserialize(ByVal lpXML As String, Optional ByRef lpErrorMessage As String = "") As Object
        Try

          Return Serializer.Deserialize.XmlString(lpXML, Me.GetType)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpXML, lpErrorMessage))
          lpErrorMessage = Helper.FormatCallStack(ex)
          Return Nothing
        End Try
      End Function

      Public Function SetPropertyValue(ByVal lpPropertyName As String,
                                       ByVal lpPropertyValue As Object,
                                       ByVal lpCreateProperty As Boolean,
                                       ByVal lpMutability As HeaderProperty.PropertyMutability,
                                       ByVal lpPropertyType As PropertyType) As Boolean

        ApplicationLogging.WriteLogEntry("Enter Header::SetPropertyValue", TraceEventType.Verbose)
        Try


          If PropertyExists(lpPropertyName) = False Then
            ' The property does not currently exist
            If lpCreateProperty = True Then
              Properties.Add(New HeaderProperty(lpPropertyType, lpPropertyName, lpPropertyValue, lpMutability))
              Return True
            End If
            Throw New InvalidOperationException(String.Format(
                                                              "Cannot set the value of '{0}' to '{1}' as the property does not exist in the header and the parameter lpCreateProperty was set to False.",
                                                              lpPropertyValue.ToString, lpPropertyName))
          Else
            ' The property current exists, just set the value
            Properties(lpPropertyName).Value = lpPropertyValue

          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        Finally
          ApplicationLogging.WriteLogEntry("Exit Header::SetPropertyValue", TraceEventType.Verbose)
        End Try

      End Function

      Public Function SetPropertyValue(ByVal lpPropertyName As String,
                                       ByVal lpPropertyValue As Object) As Boolean

        ApplicationLogging.WriteLogEntry("Enter Header::SetPropertyValue", TraceEventType.Verbose)
        Try

          Return SetPropertyValue(lpPropertyName, lpPropertyValue, False, HeaderProperty.PropertyMutability.ReadWrite, PropertyType.ecmString)

        Catch InvalidOpEx As InvalidOperationException
          If InvalidOpEx.Message.ToLower.Contains("property does not exist") Then
            Throw New InvalidOperationException(String.Format(
                                                              "Cannot set the value of '{0}' to '{1}' as the property does not exist in the header.",
                                                              lpPropertyValue.ToString, lpPropertyName))
          Else
            ApplicationLogging.LogException(InvalidOpEx, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        Finally
          ApplicationLogging.WriteLogEntry("Exit Header::SetPropertyValue", TraceEventType.Verbose)
        End Try
      End Function

      ''' <summary>
      ''' Determine if a property exists in the document.
      ''' </summary>
      ''' <param name="lpName">The name of the property to check</param>
      ''' <remarks></remarks>
      Public Function PropertyExists(ByVal lpName As String) As Boolean

        ApplicationLogging.WriteLogEntry("Enter Header::PropertyExists", TraceEventType.Verbose)
        Try

          Return Properties.PropertyExists(lpName)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ApplicationLogging.WriteLogEntry("Exit Header::PropertyExists", TraceEventType.Verbose)
        Finally
          ApplicationLogging.WriteLogEntry("Exit Header::PropertyExists", TraceEventType.Verbose)
        End Try

      End Function

      Private Sub LoadFromXmlDocument(ByVal lpXML As String)
        Try
          Dim lobjHeader As DocumentHeader = Deserialize(lpXML)
          Helper.AssignObjectProperties(lobjHeader, Me)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Helper.DumpException(ex)
        End Try
      End Sub

      Public Shadows Function ToString() As String
        Try
          Return Serializer.Serialize.XmlString(Me)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Function

      'Private Function DebuggerIdentifier() As String
      '  Dim lobjIdentifierBuilder As New StringBuilder
      '  Try

      '    Dim lstrPropertyDelimiter As String = " : "

      '    If Workstation.Length > 0 Then
      '      lobjIdentifierBuilder.AppendFormat("{0}{1}", Workstation, lstrPropertyDelimiter)
      '    End If

      '    If UserID.Length > 0 Then
      '      lobjIdentifierBuilder.AppendFormat("{0}{1}", UserID, lstrPropertyDelimiter)
      '    End If

      '    If CtsVersion.Length > 0 Then
      '      lobjIdentifierBuilder.AppendFormat("{0}{1}", CtsVersion, lstrPropertyDelimiter)
      '    End If

      '    If ID Is Nothing OrElse ID.Length = 0 Then
      '      lobjIdentifierBuilder.AppendFormat("Document identifier not set{0}", lstrPropertyDelimiter)
      '    Else
      '      lobjIdentifierBuilder.AppendFormat("{0}{1}", ID, lstrPropertyDelimiter)
      '    End If

      '    If ContentSource IsNot Nothing Then
      '      lobjIdentifierBuilder.AppendFormat("{0}", ContentSource.DebuggerIdentifier)
      '    Else
      '      lobjIdentifierBuilder.Remove(lobjIdentifierBuilder.Length - lstrPropertyDelimiter.Length, lstrPropertyDelimiter.Length)
      '    End If

      '    Return lobjIdentifierBuilder.ToString

      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return lobjIdentifierBuilder.ToString
      '  End Try
      'End Function

#Region "Constructors"

      Public Sub New()

      End Sub

      Friend Sub New(ByVal lpXML As String)
        Try
          LoadFromXmlDocument(lpXML)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Helper.DumpException(ex)
        End Try

      End Sub
      ''' <summary>
      ''' Creates a new Header using the specified DocumentID and SerializationPath
      ''' </summary>
      ''' <param name="lpDocumentID">The document identifier</param>
      ''' <param name="lpSerializationPath">The serialization path for the document</param>
      ''' <remarks>The </remarks>
      Public Sub New(ByVal lpDocumentID As String,
                     ByVal lpSerializationPath As String,
                     ByVal lpOriginalHash As String)
        Me.New(lpDocumentID, lpSerializationPath, lpOriginalHash, Nothing, Now, Environment.MachineName, Environment.UserName)
      End Sub

      ''' <summary>
      ''' Creates a new Header using the specified DocumentID, SerializationPath and ContentSource
      ''' </summary>
      ''' <param name="lpDocumentID">The document identifier</param>
      ''' <param name="lpSerializationPath">The serialization path for the document</param>
      ''' <param name="lpContentSource">The ContentSource used to export the document</param>
      ''' <remarks>The ExportDate, ExportWorkstation, and ExportUsedID will be calculated automatically</remarks>
      Public Sub New(ByVal lpDocumentID As String,
                     ByVal lpSerializationPath As String,
                     ByVal lpOriginalHash As String,
                     ByVal lpContentSource As ContentSource)
        Me.New(lpDocumentID, lpSerializationPath, lpOriginalHash, lpContentSource, Now, Environment.MachineName, Environment.UserName)
      End Sub

      ''' <summary>
      ''' Creates a new Header using the specified information
      ''' </summary>
      ''' <remarks></remarks>
      Public Sub New(ByVal lpId As String,
                     ByVal lpSerializationPath As String,
                     ByVal lpOriginalHash As String,
                     ByVal lpContentSource As Object,
                     ByVal lpGenerationDate As DateTime,
                     ByVal lpWorkstation As String,
                     ByVal lpUserID As String)

        Try
          mstrID = lpId
          mstrSerializationPath = lpSerializationPath
          mstrOriginalHash = lpOriginalHash
          'mobjContentSource = lpContentSource
          mdatGenerationDate = lpGenerationDate
          mdatGenenerationUtcDate = lpGenerationDate.ToUniversalTime
          mstrWorkstation = lpWorkstation
          mstrUserID = lpUserID
          mstrCtsVersion = Assembly.GetExecutingAssembly.GetName.Version.ToString() 'FrameworkVersion.CurrentVersion
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

      End Sub

#End Region

#Region "Private Methods"

      ''' <summary>
      ''' Will return the latest hash of the document, 
      ''' whether or not it was the original hash or the 
      ''' hash of the last transformation performed.
      ''' </summary>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Friend Function GetLatestHash() As String
        Try

          If TransformationSeries.Count = 0 Then
            ' Return the original document hash
            Return Me.OriginalHash
          Else
            ' Get the hash of the last transformation performed
            Return TransformationSeries(TransformationSeries.Count - 1).Hash
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

#End Region

    End Class

    ''' <summary>
    ''' Encapsulates information related to a transformation executed on the document
    ''' </summary>
    ''' <remarks></remarks>
    ''' <exclude/>
    <Microsoft.VisualBasic.HideModuleName()>
    <Serializable()>
    Public Class Modification

#Region "Class Variables"

      Private mstrHash As String = ""
      Private mdatTransformDate As DateTime
      Private mstrWorkstation As String = ""
      Private mstrUserID As String = ""
      Private mobjTransformation As Transformation

#End Region

#Region "Public Properties"

      ''' <summary>
      ''' The hash of the document immediately following this transformation
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks>Used to verify whether or not the document was modified subsequent to this transformation</remarks>
      Public Property Hash() As String
        Get
          Return mstrHash
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrHash = value
            Else
              Throw New InvalidOperationException(String.Format("Although {0} is a public property, add operations are not allowed.  Treat property as read-only.", Reflection.MethodBase.GetCurrentMethod.Name))
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The date and time the transformation was executed against the document
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property TransformDate() As DateTime
        Get
          Return mdatTransformDate
        End Get
        Set(ByVal value As DateTime)
          Try
            If Helper.IsDeserializationBasedCall Then
              mdatTransformDate = value
            Else
              Throw New InvalidOperationException(String.Format("Although {0} is a public property, add operations are not allowed.  Treat property as read-only.", Reflection.MethodBase.GetCurrentMethod.Name))
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The workstation used to execute the transformation
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property Workstation() As String
        Get
          Return mstrWorkstation
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrWorkstation = value
            Else
              Throw New InvalidOperationException(String.Format("Although {0} is a public property, add operations are not allowed.  Treat property as read-only.", Reflection.MethodBase.GetCurrentMethod.Name))
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The network user id of the person that transformed the document
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property UserID() As String
        Get
          Return mstrUserID
        End Get
        Set(ByVal value As String)
          Try
            If Helper.IsDeserializationBasedCall Then
              mstrUserID = value
            Else
              Throw New InvalidOperationException(String.Format("Although {0} is a public property, add operations are not allowed.  Treat property as read-only.", Reflection.MethodBase.GetCurrentMethod.Name))
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

      ''' <summary>
      ''' The transformation used to transform the document
      ''' </summary>
      ''' <value></value>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Property Transformation() As Transformation
        Get
          Return mobjTransformation
        End Get
        Set(ByVal value As Transformation)
          Try
            If Helper.IsDeserializationBasedCall Then
              mobjTransformation = value
            Else
              Throw New InvalidOperationException(String.Format("Although {0} is a public property, add operations are not allowed.  Treat property as read-only.", Reflection.MethodBase.GetCurrentMethod.Name))
            End If
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
        End Set
      End Property

#End Region

#Region "Constructors"

      Public Sub New()
        'Me.New("", "", Nothing)
      End Sub

      Public Sub New(ByVal lpWorkstation As String,
                     ByVal lpUserID As String,
                     ByVal lpTransformation As Transformation)
        mstrWorkstation = lpWorkstation
        mstrUserID = lpUserID
        mobjTransformation = lpTransformation
        mdatTransformDate = Now.ToUniversalTime

        mstrHash = lpTransformation.Document.GenerateHash

      End Sub

#End Region

    End Class

    ''' <exclude/>
    <Microsoft.VisualBasic.HideModuleName()>
    <Serializable()>
    Public Class Modifications
      Inherits CCollection(Of Modification)

      Public Overloads Sub Add(ByVal lpIteration As Modification)
        Try
          If Helper.IsDeserializationBasedCall OrElse Helper.CallStackContainsMethodName("TransformDocument") Then
            MyBase.Add(lpIteration)
          Else
            Throw New InvalidOperationException("Although this is a public method, add operations are not allowed.  Treat collection as read-only.")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Sub

      Public Sub New()

      End Sub
    End Class

#End Region

#Region "Class Variables"

    <NonSerializedAttribute()>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjContentSource As ContentSource

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Protected mobjRMProvider As IBasicContentServicesProvider

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjHeader As DocumentHeader

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjDataTable As DataTable

    Private WithEvents MobjFolderPaths As ObservableCollection(Of String)

#End Region

#Region "Public Properties"

    <JsonIgnore()>
    Public ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

    ' ''' <summary>
    ' ''' Gets or sets the parent and/or child relationships for the document.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property Relationships() As Relationships
    '  Get

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return mobjRelationships

    '  End Get
    '  Set(ByVal value As Relationships)

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    mobjRelationships = value

    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' Gets or sets the collection of properties for the document.
    ' ''' </summary>
    'Public Property Properties() As ECMProperties
    '  Get

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return mobjProperties

    '  End Get
    '  Set(ByVal value As ECMProperties)

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    mobjProperties = value

    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' Gets or sets the collection of metadata properties for the object.
    ' ''' </summary>
    ' ''' <returns>ECMProperties collection</returns>
    ' ''' <remarks>Interface passthrough to the Properties collection.</remarks>
    '<XmlIgnore()> _
    'Public Property Metadata() As ECMProperties Implements IMetaHolder.Metadata
    '  Get

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return Properties

    '  End Get
    '  Set(ByVal value As ECMProperties)

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Properties = value

    '  End Set
    'End Property


    ' ''' <summary>
    ' ''' Gets or sets the class of document to which this document belongs.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property DocumentClass() As String
    '  Get
    '    Try

    '      If IsDisposed Then
    '        Throw New ObjectDisposedException(Me.GetType.ToString)
    '      End If

    '      If Properties.PropertyExists("Document Class") = False Then
    '        'Properties.Add(New ECMProperty(PropertyType.ecmString, _
    '        '                               "Document Class", ""))
    '        Properties.Add(PropertyFactory.Create(PropertyType.ecmString, _
    '                                              "Document Class", String.Empty))
    '      End If
    '      Return Properties("Document Class").Value.ToString
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      Return String.Empty
    '    End Try
    '  End Get

    '  Set(ByVal Value As String)
    '    Try

    '      If IsDisposed Then
    '        Throw New ObjectDisposedException(Me.GetType.ToString)
    '      End If

    '      If Properties.PropertyExists("Document Class") Then
    '        Properties("Document Class").Value = Value
    '      Else
    '        'Properties.Add(New ECMProperty(PropertyType.ecmString, _
    '        '                               "Document Class", Value))
    '        Properties.Add(PropertyFactory.Create(PropertyType.ecmString, _
    '                                  "Document Class", Value))

    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("Document::Set_DocumentClass('{0}')", Value))
    '    End Try
    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' The object identifier for the document.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks>This value is typically initialized to the new document 
    ' ''' identifier during a document add or import operation.</remarks>
    'Public Property ObjectID() As String
    '  Get
    '    Try

    '      If IsDisposed Then
    '        Throw New ObjectDisposedException(Me.GetType.ToString)
    '      End If

    '      If Properties.PropertyExists("ObjectID") = False Then
    '        'Properties.Add(New ECMProperty(PropertyType.ecmString, _
    '        '                               "ObjectID", ""))
    '        Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "ObjectID", String.Empty))
    '      End If
    '      Return Properties("ObjectID").Value.ToString
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      Return ""
    '    End Try
    '  End Get
    '  Set(ByVal Value As String)
    '    Try

    '      If IsDisposed Then
    '        Throw New ObjectDisposedException(Me.GetType.ToString)
    '      End If

    '      If Properties.PropertyExists("ObjectID") Then
    '        Properties("ObjectID").Value = Value
    '      Else
    '        'Properties.Add(New ECMProperty(PropertyType.ecmString, _
    '        '                               "ObjectID", Value))
    '        Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "ObjectID", Value))
    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("Value::Set_ObjectID('{0}')", Value))
    '    End Try
    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' The path the document was serialized to.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    '<XmlIgnore()> _
    'Public Property SerializationPath() As String
    '  ' Intended to be used for re-writing out file after transformation
    '  Get

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return mstrSerializationPath

    '  End Get
    '  Set(ByVal value As String)

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    mstrSerializationPath = value

    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' The path the document was deserialized to.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    '<XmlIgnore()> _
    'Public Property DeSerializationPath() As String
    '  ' Intended to be used for re-writing out file after transformation
    '  Get

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return mstrDeSerializationPath

    '  End Get
    '  Set(ByVal value As String)

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    mstrDeSerializationPath = value

    '  End Set
    'End Property

    ''' <summary>
    ''' Gets the folder location the document was serialized to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property DocumentPath() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          Return GetDocumentPath()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the collection of folder names in which this document is filed.
    ''' </summary>
    ''' <param name="lpAddBasePath"></param>
    ''' <param name="lpAddPathLocation"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public ReadOnly Property FolderPaths(Optional ByVal lpAddBasePath As String = "",
                                         Optional ByVal lpAddPathLocation As ePathLocation = ePathLocation.Front) As ObservableCollection(Of String) 'Collections.Specialized.StringCollection
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          If MobjFolderPaths Is Nothing Then
            MobjFolderPaths = GetFolderPaths(lpAddBasePath)
          End If
          Return MobjFolderPaths
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ' ''' <summary>
    ' ''' The name of the document.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks>If no property is defined for 'Name' 
    ' ''' this may also return the value of the following 
    ' ''' properties respectively, 
    ' ''' ('Title', 'DocumentTitle', 'Document Title', 'FileName', 'Subject')</remarks>
    'Public ReadOnly Property Name() As String
    '  Get
    '    Try

    '      If IsDisposed Then
    '        Throw New ObjectDisposedException(Me.GetType.ToString)
    '      End If

    '      Dim lobjNameValue As Object
    '      If Me.PropertyExists(PropertyScope.DocumentProperty, "Name") Then
    '        lobjNameValue = Properties("Name").Value
    '      ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Title") Then
    '        lobjNameValue = GetProperty("Title", PropertyScope.VersionProperty).Value
    '      ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "DocumentTitle") Then
    '        lobjNameValue = GetProperty("DocumentTitle", PropertyScope.VersionProperty).Value 'Properties("DocumentTitle").Value
    '      ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Document Title") Then
    '        lobjNameValue = GetProperty("Document Title", PropertyScope.VersionProperty).Value 'Properties("DocumentTitle").Value
    '      ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "FileName") Then
    '        lobjNameValue = GetProperty("FileName", PropertyScope.VersionProperty).Value 'Properties("Subject").Value
    '      ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Subject") Then
    '        lobjNameValue = GetProperty("Subject", PropertyScope.VersionProperty).Value 'Properties("Subject").Value

    '      ElseIf mobjID IsNot Nothing Then
    '        lobjNameValue = New Value(ID)
    '      Else
    '        lobjNameValue = New Value("")
    '      End If

    '      If (lobjNameValue.GetType().IsInstanceOfType(GetType(Value))) Then
    '        Return lobjNameValue.Value.ToString
    '      ElseIf lobjNameValue.ToString = "Ecmg.Cts.Core.Value" Then
    '        Return lobjNameValue.value
    '      Else
    '        Return lobjNameValue.ToString()
    '      End If


    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      Return ""
    '    End Try
    '  End Get
    'End Property

    'Public ReadOnly Property FolderPathArray(Optional ByVal lpAddBasePath As String = "", _
    '  Optional ByVal lpAddPathLocation As ePathLocation = ePathLocation.Front, _
    '  Optional ByVal lpDelimiter As String = "/", _
    '  Optional ByVal lpLeadingDelimiter As Boolean = True) As String()
    '  Get
    '    Return GetFolderPathArray(lpAddBasePath, lpAddPathLocation, _
    '      lpDelimiter, lpLeadingDelimiter)
    '  End Get
    'End Property

    ''' <summary>
    ''' Returns the array of folders for the specified PathFactory
    ''' </summary>
    ''' <param name="lpPathFactory"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public ReadOnly Property FolderPathArray(ByVal lpPathFactory As PathFactory) As String()
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          Return GetFolderPathArray(lpPathFactory)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    '''     Gets the default file name.
    ''' </summary>
    ''' <value>
    '''     <para>
    '''         Returns a proposed file name for the document package for serialization.  
    ''' First it will use the ID property, if that is empty it will use the Name property.
    '''     </para>
    ''' </value>
    ''' <remarks>
    '''     If there are any illegal file system characters in the Id or Name 
    '''     property they will be replaced with an '_' character.
    ''' </remarks>
    <JsonIgnore()>
    Public ReadOnly Property DefaultFileName As String
      Get
        Try
          If Not String.IsNullOrEmpty(Me.ID) Then
            Return Helper.CleanFile(String.Format("{0}.cpf", Me.ID), "_")
          ElseIf Not String.IsNullOrEmpty(Me.Name) Then
            Return Helper.CleanFile(String.Format("{0}.cpf", Me.Name), "_")
          Else
            Return "document.cpf"
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' The hash of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Used to determine whether or not the document has been modified since originally created.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property Hash() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          Return mstrHash
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          ' This can only be set internally as it is intended to be used as a check for changed data
          If Helper.IsDeserializationBasedCall Then
            mstrHash = value
          End If
          '' Create a StackTrace object with file names and line numbers
          'Dim lobjStackTrace As New System.Diagnostics.StackTrace()
          'Dim lstrCallingMethod As String

          'For i As Integer = 0 To lobjStackTrace.FrameCount - 1
          '  With lobjStackTrace.GetFrame(i)
          '    lstrCallingMethod = .GetMethod.ToString.ToLower

          '    If lstrCallingMethod.Contains("deserialize") OrElse lstrCallingMethod.Contains("loadfromxmldocument") Then
          '      mstrHash = value
          '      Exit Property
          '    End If
          '  End With
          'Next i
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The header object for the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property Header() As DocumentHeader
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          'If mobjHeader Is Nothing Then
          '  UpdateHeader()
          'End If
          Return mobjHeader
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Used to perform repository operations on the document.  
    ''' </summary>
    ''' <value></value>
    ''' <returns>A ContentSource object reference</returns>
    ''' <remarks>If set to a valid content source object in which the document resides, 
    ''' and the provider associated with the content source implements the "IBasicContentServicesProvider" interface, 
    ''' the document can perform the operations exposed by that interface.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property ContentSource() As ContentSource
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          Return mobjContentSource
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As ContentSource)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

          mobjContentSource = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    '' ''' <summary>
    '' ''' Gets or sets the type of storage associated with the document content
    '' ''' </summary>
    '' ''' <value>The StorageTypeEnum value to set for the Document StorageType</value>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public Property StorageType() As Content.StorageTypeEnum
    ''  Get


    ''    If IsDisposed Then
    ''      Throw New ObjectDisposedException(Me.GetType.ToString)
    ''    End If

    ''    Return menuStorageType

    ''  End Get

    ''  Set(ByVal value As Content.StorageTypeEnum)


    ''    If IsDisposed Then
    ''      Throw New ObjectDisposedException(Me.GetType.ToString)
    ''    End If

    ''    menuStorageType = value

    ''    For Each lobjVersion As Version In Versions
    ''      For Each lobjContent As Content In lobjVersion.Contents
    ''        lobjContent.StorageType = value
    ''      Next
    ''    Next

    ''  End Set
    ''End Property

    '' ''' <summary>
    '' ''' Gets the first version of the document
    '' ''' </summary>
    '' ''' <value></value>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public ReadOnly Property FirstVersion As Version
    ''  Get
    ''    Return GetFirstVersion()
    ''  End Get
    ''End Property

    '' ''' <summary>
    '' ''' Gets the latest version of the document
    '' ''' </summary>
    '' ''' <value></value>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public ReadOnly Property LatestVersion As Version
    ''  Get
    ''    Return GetLatestVersion()
    ''  End Get
    ''End Property

    ''' <summary>
    ''' Gets the total number of content elements from all versions of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property ContentCount As Integer
      Get
        Try
          Return GetContentCount()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new document and associates it with the specified content source.
    ''' </summary>
    ''' <param name="lpSourceContentSource">
    ''' The content source used to create or export the document.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpSourceContentSource As ContentSource)
      MyBase.New()
      Try
        Me.ContentSource = lpSourceContentSource
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new document and associates it with the specified source provider.
    ''' </summary>
    ''' <param name="lpSourceProvider">
    ''' The provider used to create or export the document.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpSourceProvider As IProvider)
      Try
        Me.ContentSource = lpSourceProvider.ContentSource
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' ''' <summary>
    ' ''' Constructs a new Document object from the specified stream.
    ' ''' </summary>
    ' ''' <param name="lpStream">The stream to create the document from.</param>
    ' ''' <remarks>Will attempt to create the document from 
    ' ''' either a native stream or an archive stream.</remarks>
    'Public Sub New(ByVal lpStream As Stream)
    '  Try

    '    Dim lobjDocument As Document

    '    Dim lobjSeekableStream As Stream = lpStream

    '    ' If we cannot seek, then create a copy of the stream
    '    ' in memory and pass that along
    '    If (lobjSeekableStream.CanSeek = False) Then

    '      lobjSeekableStream = New MemoryStream()
    '      Dim buf(1023) As Byte
    '      Dim numRead As Integer = lpStream.Read(buf, 0, 1024)
    '      While (numRead > 0)
    '        lobjSeekableStream.Write(buf, 0, numRead)
    '        numRead = lpStream.Read(buf, 0, 1024)
    '      End While

    '      lobjSeekableStream.Position = 0

    '    End If

    '    ' Check to see if the stream is for an xml file or not
    '    If Helper.IsXmlStream(lobjSeekableStream) Then
    '      ' The stream is a native xml file, this would be for a cdf file.
    '      lobjDocument = CreateFromStream(lobjSeekableStream)
    '    ElseIf Helper.IsZipStream(lobjSeekableStream) Then
    '      ' The stream is not a native xml file, this would most likely be for a cpf file.
    '      lobjDocument = FromArchive(lobjSeekableStream)
    '    Else
    '      Throw New DocumentException(String.Empty, "An xml stream or a zip stream was expected.")
    '    End If

    '    Helper.AssignObjectProperties(lobjDocument, Me)

    '    SetDocumentReferences()

    '    ' We are checking to see if the existing 'cdf' file has a header
    '    ' If it does not have a header we want to skip over the update header
    '    If lobjDocument.HeaderString.Length > 0 Then
    '      '  Recreate the header
    '      UpdateHeader()
    '    End If

    '    mstrCurrentPath = lobjDocument.CurrentPath

    '  Catch BadPasswordEx As BadPasswordException
    '    Dim lobjDocumentEx As New DocumentException(String.Empty, "A password was expected but not supplied.", BadPasswordEx)
    '    ApplicationLogging.LogException(lobjDocumentEx, Reflection.MethodBase.GetCurrentMethod)
    '    Throw lobjDocumentEx
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ' ''' <summary>
    ' ''' Constructs a new Document object from the specified archive stream and password.
    ' ''' </summary>
    ' ''' <param name="lpStream">The archive stream to create the document from.</param>
    ' ''' <param name="lpPassword">The password to unlock the archive stream.</param>
    ' ''' <remarks></remarks>
    'Public Sub New(ByVal lpStream As Stream, ByVal lpPassword As String)
    '  Try

    '    Dim lobjDocument As Document = FromArchive(lpStream, lpPassword)

    '    Helper.AssignObjectProperties(lobjDocument, Me)

    '    SetDocumentReferences()

    '    ' We are checking to see if the existing 'cdf' file has a header
    '    ' If it does not have a header we want to skip over the update header
    '    If lobjDocument.HeaderString.Length > 0 Then
    '      '  Recreate the header
    '      UpdateHeader()
    '    End If

    '    mstrCurrentPath = lobjDocument.CurrentPath

    '  Catch BadPasswordEx As BadPasswordException
    '    Dim lobjDocumentEx As New DocumentException(String.Empty, "The specified password does not match.")
    '    ApplicationLogging.LogException(lobjDocumentEx, Reflection.MethodBase.GetCurrentMethod)
    '    Throw lobjDocumentEx
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ''' <summary>
    ''' Creates a document from a saved file.  The file can be either a Content Package File (*.cpf) 
    ''' or a Content Definition File (*.cdf).
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified path to the document file.</param>
    ''' <remarks></remarks>
    ''' <example>
    '''  <para>Below are some examples of creating Document objects from saved files.</para>
    '''  <para></para>
    '''  <para></para>
    '''  <code title="CDF Example" description="In this example we will build a Document object from a cdf file." lang="VB.NET">
    ''' Dim lobjDocument as New Document("C:\Temp\3849872.cdf")</code>
    '''  <code title="CPF Example" description="In this example we will create a Document object from a cpf file." lang="VB.NET">
    ''' Dim lobjDocument as New Document("C:\Temp\3849872.cpf")</code>
    ''' </example>
    Public Sub New(ByVal lpFilePath As String)
      Try

        Dim lstrErrorMessage As String = String.Empty
        Dim lobjDocument As Document

        If lpFilePath Is Nothing OrElse lpFilePath.Length = 0 Then
          Throw New InvalidPathException
        End If

        If String.Equals(Path.GetExtension(lpFilePath).Replace(".", String.Empty),
                         CONTENT_PACKAGE_FILE_EXTENSION, StringComparison.InvariantCultureIgnoreCase) Then
          lobjDocument = FromArchive(lpFilePath)
        ElseIf String.Equals(Path.GetExtension(lpFilePath).Replace(".", String.Empty),
                         JSON_CONTENT_PACKAGE_FILE_EXTENSION, StringComparison.InvariantCultureIgnoreCase) Then
          lobjDocument = FromArchive(lpFilePath)
        Else
          lobjDocument = Deserialize(lpFilePath, lstrErrorMessage)
        End If



        If lobjDocument Is Nothing Then
          ' Check the error message
          If lstrErrorMessage.Length > 0 Then
            Throw New Exception(lstrErrorMessage)
          End If
        End If

        Helper.AssignObjectProperties(lobjDocument, Me)

        VerifyProperties(lpFilePath)

        SetDocumentReferences()

        ' We are checking to see if the existing 'cdf' file has a header
        ' If it does not have a header we want to skip over the update header
        If lobjDocument.HeaderString IsNot Nothing AndAlso lobjDocument.HeaderString.Length > 0 Then
          '  Recreate the header
          UpdateHeader()
        End If

        mstrCurrentPath = lpFilePath

      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'LogSession.LogException(ex)
        'Helper.DumpException(ex)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      Try

        LoadFromXmlDocument(lpXML)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "ISerialize Implementation"

    'Private Sub PrepareContentForSerialization()
    '  Try

    '    Select Case Me.StorageType
    '      Case Content.StorageTypeEnum.EncodedCompressed

    '    End Select
    '    '  Encode the file into the cdf
    '    EncodeAllContents(True)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return CONTENT_DEFINITION_FILE_EXTENSION
      End Get
    End Property

    ''' <summary>
    ''' Instantiate from an XML file.
    ''' </summary>
    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lobjDocument As Document = Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
        'lobjDocument.CurrentPath = lpFilePath
        'lobjDocument.SerializationPath = lpFilePath
        lobjDocument.DeSerializationPath = lpFilePath
        Return lobjDocument
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.DeSerialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Serializes the document to a .cdf file
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified output path</param>
    ''' <param name="lpZip">Specifies whether or not the resulting file should be zipped</param>
    ''' <remarks>If lpZip is set to True then the resulting file will be a zip file containing a single cdf file.</remarks>
    <Obsolete("While still supported, this method is obsolete, consider using Save or Archive instead")>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpZip As Boolean)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Save(lpFilePath)

        If lpZip = True Then

          '  We were asked to zip up the file

          '  Create the file name for the new zip file
          Dim lstrZipFilePath As String = String.Format("{0}.zip", lpFilePath)

          '  If this file already exists we should delete it
          If IO.File.Exists(lstrZipFilePath) Then
            IO.File.Delete(lstrZipFilePath)
          End If

          '  Create the zip file
          Dim lobjZipFile As New ZipFile(lstrZipFilePath)

          '  Add the cdf file to the zip
          lobjZipFile.AddFile(lpFilePath)

          '  Save the zip file
          lobjZipFile.Save()

          '  Delete the original cdf file
          IO.File.Delete(lpFilePath)

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the Document object in an XML based cdf file.
    ''' </summary>
    <Obsolete("While still supported, this method is obsolete, consider using Save instead")>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Save(lpFilePath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <Obsolete("While still supported, this method is obsolete, consider using Save instead")>
    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpContentReferenceBehavior As ContentReferenceBehavior)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Save(lpFilePath, lpContentReferenceBehavior)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the Document object in an XML file.
    ''' </summary>
    ''' <param name="lpFileExtension">Used if desired to override the default file extension of .cdf</param>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Save(lpFilePath, lpFileExtension, False, "")

        mstrCurrentPath = lpFilePath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Serializes the document to a .cdf file
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified output path</param>
    ''' <param name="lpWriteProcessingInstruction">Specifies whether or not an XML write procesing instruction is to be provided</param>
    ''' <param name="lpStyleSheetPath">If the value of lpWriteProcessingInstruction is set to true, then the specified style sheet path is inserted into the file</param>
    ''' <remarks></remarks>
    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Save(lpFilePath, "", lpWriteProcessingInstruction, lpStyleSheetPath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Generates an XmlDocument representation of the current state of the document
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        SetDocumentReferences()
        '  Encode the contents as appropriate for the storage type
        EncodeAllContents()
        UpdateHeader()

        Return Serializer.Serialize.Xml(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Shared method to create a document from a stream
    ''' </summary>
    ''' <param name="lpStream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateFromStream(ByVal lpStream As IO.Stream) As Document

      Try

        If (lpStream.CanSeek) Then
          lpStream.Position = 0
        End If

        Return Serializer.Deserialize.FromStream(lpStream, GetType(Document))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    '' ''' <summary>
    '' ''' Creates and returns a stream of the current document.
    '' ''' </summary>
    '' ''' <returns></returns>
    '' ''' <remarks>
    '' ''' If you want the stream to include all of the 
    '' ''' document content and the storage type is not 
    '' ''' encoded, use ToArchiveStream instead.
    '' ''' </remarks>
    ''Public Function ToStream() As IO.MemoryStream

    ''  Try

    ''    If IsDisposed Then
    ''      Throw New ObjectDisposedException(Me.GetType.ToString)
    ''    End If

    ''    SetDocumentReferences()

    ''    '  Encode the contents as appropriate for the storage type
    ''    EncodeAllContents()

    ''    Return Serializer.Serialize.ToStream(Me)

    ''    'Return ToArchiveStream()

    ''  Catch ex As Exception
    ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''    '  Re-throw the exception to the caller
    ''    Throw
    ''  End Try

    ''End Function

    ' ''' <summary>
    ' ''' Creates and returns an ArchiveStream using the specified password.
    ' ''' </summary>
    ' ''' <param name="password"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function ToStream(ByVal password As String) As IO.MemoryStream
    '  Try
    '    Return ToArchiveStream(password)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ''' <summary>
    ''' Generates an XML string representation of the current state of the document
    ''' </summary>
    ''' <returns>XML string</returns>
    ''' <remarks>Overides Object.ToString</remarks>
    Public Function ToXmlString() As String Implements ISerialize.ToXmlString

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (Me.StorageType = Content.StorageTypeEnum.Reference) Then
          SetDocumentReferences()
          '  Encode the contents as appropriate for the storage type
          EncodeAllContents()
          'InitializeHash()
          Return Serializer.Serialize.XmlString(Me)
        Else
          Return Me.ID
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlElementString() As String
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (Me.StorageType = Content.StorageTypeEnum.Reference) Then
          SetDocumentReferences()
          '  Encode the contents as appropriate for the storage type
          EncodeAllContents()
          'InitializeHash()
          Return Serializer.Serialize.XmlElementString(Me)
        Else
          Return Me.ID
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

#Region "Document Property Methods"

    ''' <summary>
    ''' Gets a specific property using a LookupProperty reference.
    ''' </summary>
    ''' <param name="lpLookupProperty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProperty(ByVal lpLookupProperty As Transformations.LookupProperty) As ECMProperty
      Try
        Return GetProperty(lpLookupProperty.PropertyName, lpLookupProperty.PropertyScope, lpLookupProperty.VersionIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Finds the specified property in the document and returns the property scope of the property.
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <returns></returns>
    ''' <remarks>If the property is not defined in the document a PropertyDoesNotExistException exception is thrown.</remarks>
    Public Function GetPropertyScope(ByVal lpProperty As ECMProperty) As PropertyScope
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpProperty, "The lpProperty parameter is null, a valid ECMProperty reference is required for this method!")
#Else
          If lpProperty Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpProperty), "The lpProperty parameter is null, a valid ECMProperty reference is required for this method!")
          End If
#End If

        If Properties.PropertyExists(lpProperty.Name) Then
          Return PropertyScope.DocumentProperty
        ElseIf GetLatestVersion.Properties.PropertyExists(lpProperty.Name) Then
          Return PropertyScope.VersionProperty
        Else
          Throw New PropertyDoesNotExistException("Property '" & lpProperty.Name & "' does not exist in the document.", lpProperty.Name)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' ''' <summary>
    ' ''' Gets a specific property by name
    ' ''' </summary>
    ' ''' <param name="lpPropertyName"></param>
    ' ''' <param name="lpVersionIndex"></param>
    ' ''' <returns></returns>
    ' ''' <remarks>Checks for the property first at the docuemnt level, then at the version level.</remarks>
    'Public Function GetProperty(ByVal lpPropertyName As String, _
    '                            Optional ByVal lpVersionIndex As Integer = 0) As ECMProperty
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return GetProperty(lpPropertyName, PropertyScope.BothDocumentAndVersionProperties, lpVersionIndex)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ' ''' <summary>
    ' ''' Gets a specific property by name
    ' ''' </summary>
    ' ''' <param name="lpPropertyName"></param>
    ' ''' <param name="lpPropertyScope"></param>
    ' ''' <param name="lpVersionIndex"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function GetProperty(ByVal lpPropertyName As String, _
    '                            ByVal lpPropertyScope As PropertyScope, _
    '                            Optional ByVal lpVersionIndex As Integer = 0) As ECMProperty

    '  ApplicationLogging.WriteLogEntry("Enter Document::GetProperty", TraceEventType.Verbose)
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Select Case lpPropertyScope
    '      Case PropertyScope.DocumentProperty
    '        Return Properties(lpPropertyName)

    '      Case PropertyScope.VersionProperty
    '        If lpVersionIndex = ALL_VERSIONS Then
    '          Return GetLatestVersion.Properties(lpPropertyName)
    '        Else
    '          Return Versions(lpVersionIndex).Properties(lpPropertyName)
    '        End If

    '      Case PropertyScope.BothDocumentAndVersionProperties
    '        If Properties.PropertyExists(lpPropertyName) Then
    '          Return Properties(lpPropertyName)
    '        ElseIf Versions(lpVersionIndex).Properties.PropertyExists(lpPropertyName) Then
    '          Return Versions(lpVersionIndex).Properties(lpPropertyName)
    '        Else
    '          Throw New PropertyDoesNotExistException("Property '" & lpPropertyName & "' does not exist in the document.", lpPropertyName)
    '        End If

    '      Case PropertyScope.ContentProperty
    '        If lpVersionIndex = ALL_VERSIONS Then
    '          'Return GetLatestVersion.GetPrimaryContent.Metadata.Item(lpPropertyName)
    '          Return GetLatestVersion.GetPrimaryContent.CompleteProperties(lpPropertyName)
    '        Else
    '          'Return Versions(lpVersionIndex).GetPrimaryContent.Metadata.Item(lpPropertyName)
    '          Return Versions(lpVersionIndex).GetPrimaryContent.CompleteProperties(lpPropertyName)
    '        End If

    '      Case PropertyScope.AllProperties
    '        If Properties.PropertyExists(lpPropertyName) Then
    '          Return Properties(lpPropertyName)
    '        ElseIf Versions(lpVersionIndex).Properties.PropertyExists(lpPropertyName) Then
    '          Return Versions(lpVersionIndex).Properties(lpPropertyName)
    '        ElseIf GetLatestVersion.HasContent = True AndAlso GetLatestVersion.GetPrimaryContent.Metadata.PropertyExists(lpPropertyName) Then
    '          Return GetLatestVersion.GetPrimaryContent.Metadata.Item(lpPropertyName)
    '        Else
    '          Throw New PropertyDoesNotExistException("Property '" & lpPropertyName & "' does not exist in the document.", lpPropertyName)
    '        End If

    '      Case Else
    '        Return Properties(lpPropertyName)

    '    End Select
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Throw New InvalidOperationException("Unable to Get Property '" & lpPropertyName & "'", ex)
    '  Finally
    '    ApplicationLogging.WriteLogEntry("Exit Document::GetProperty", TraceEventType.Verbose)
    '  End Try

    'End Function

    ''' <summary>
    ''' Adds a value to the specifed property in the document
    ''' </summary>
    ''' <param name="lpPropertyScope">Specifies whether the property to be changed is at the document level or the version level</param>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <param name="lpVersionIndex">The specific version on which to change the value</param>
    ''' <param name="lpAllowDuplicates">Specifies whether or not to allow duplicate values to be added.</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>Only valid for multi-valued properties.
    ''' 
    ''' If the property scope is set to the version level, 
    ''' only the version specified by the lpVersionIndex will be affected.</remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If lpAllowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Function AddPropertyValue(ByVal lpPropertyScope As PropertyScope,
                                        ByVal lpName As String,
                                        ByVal lpNewValue As Object,
                                        ByVal lpVersionIndex As Integer,
                                        ByVal lpAllowDuplicates As Boolean) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Document::AddPropertyValue", TraceEventType.Verbose)

      Try


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lstrErrorMessage As String = String.Empty

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty

            If Properties(lpName).Cardinality <> Cardinality.ecmMultiValued Then
              lstrErrorMessage = String.Format("Unable to AddPropertyValue, the property '{0}' is not a multi-valued property, this action is only valid for multi-valued properties.",
                                                            lpName)
              Dim lobjDocOpEx As New InvalidOperationException(lstrErrorMessage)
              Throw New DocumentException(Me.ID, lstrErrorMessage, lobjDocOpEx)
            End If

            Properties(lpName).Values.Add(lpNewValue, lpAllowDuplicates)

          Case PropertyScope.VersionProperty
            Try

              If GetFirstVersion.Properties(lpName).Cardinality <> Cardinality.ecmMultiValued Then
                lstrErrorMessage = String.Format("Unable to AddPropertyValue, the property '{0}' is not a multi-valued property, this action is only valid for multi-valued properties.",
                                                              lpName)
                Dim lobjVerOpEx As New InvalidOperationException(lstrErrorMessage)
                Throw New DocumentException(Me.ID, lstrErrorMessage, lobjVerOpEx)
              End If

              If lpVersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
                ' This change will apply to all version of the document
                For Each lobjVersion As Version In Versions
                  lobjVersion.Properties(lpName).Values.Add(lpNewValue, lpAllowDuplicates)
                Next

              Else
                ' This change will apply to the specified version of the document
                Versions(lpVersionIndex).Properties(lpName).Values.Add(lpNewValue, lpAllowDuplicates)

              End If
            Catch ValueExistsEx As ValueExistsException
              Throw New ValueExistsException(ValueExistsEx.Value,
                String.Format("Value '{0}' not added to property '{1}' in document '{2}' because it already exists and AllowDuplicates was set to false.",
                              ValueExistsEx.Value, lpName, Me.ID))
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              ApplicationLogging.WriteLogEntry("Exit Document::ChangePropertyValue", TraceEventType.Verbose)
              Throw New InvalidOperationException("Could not add property value [" & ex.Message & "]", ex)
            End Try

        End Select

        ApplicationLogging.WriteLogEntry("Exit Document::AddPropertyValue", TraceEventType.Verbose)

        Return True

      Catch ValueExistsEx As ValueExistsException
        ' Note it as a warning in the log and pass it on.
        ApplicationLogging.WriteLogEntry(ValueExistsEx.Message, Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 23876)
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Changes the file name of the specified content element.
    ''' </summary>
    ''' <param name="lpNewName">The new name without file extension to assign.</param>
    ''' <param name="lpVersionIndex">The version index to apply the change to.</param>
    ''' <param name="lpContentElementIndex">The content element index to apply the change to.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function ChangeContentRetrievalName(ByVal lpNewName As String, ByVal lpVersionIndex As Integer, ByVal lpContentElementIndex As Integer) As Boolean

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If lpVersionIndex = ALL_VERSIONS Then
          ' This change will apply to all version of the document
          Dim lblnReturnValue As Boolean = True

          For Each lobjVersion As Version In Versions
            lblnReturnValue = lobjVersion.Contents(lpContentElementIndex).Rename(lpNewName)
            If lblnReturnValue = False Then
              Return False
            End If
          Next
        Else
          ' This change will apply to the specified version of the document
          Return Versions(lpVersionIndex).Contents(lpContentElementIndex).Rename(lpNewName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Changes the mime type for the specified content element.
    ''' </summary>
    ''' <param name="lpNewMimeType">The new mime type to assign.  If the new mime type 
    ''' is the same as the current file extension, the change will be declined.</param>
    ''' <param name="lpVersionIndex">The version index to apply the change to.</param>
    ''' <param name="lpContentElementIndex">The content element index to apply the change to.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks>Does not change any registry values, only the Cts document metadata is affected.</remarks>
    Public Overridable Function ChangeContentMimeType(ByVal lpNewMimeType As String,
                                                      ByVal lpVersionIndex As Integer,
                                                      ByVal lpContentElementIndex As Integer,
                                                      ByRef lpMessage As String) As Boolean

      Try


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lblnReturnValue As Boolean = True

        If lpVersionIndex = ALL_VERSIONS Then
          ' This change will apply to all version of the document
          For Each lobjVersion As Version In Versions
            If lobjVersion.Contents.Count = 0 Then
              Throw New InvalidOperationException("The version has no content element")
            End If
            With lobjVersion.Contents(lpContentElementIndex)
              If String.Equals(.FileExtension, lpNewMimeType, StringComparison.InvariantCultureIgnoreCase) Then
                ' The proposed new mime type is the same as the current extension
                ' Do not make this change
                lblnReturnValue = False
              Else
                .MIMEType = lpNewMimeType
                lblnReturnValue = True
              End If
            End With
          Next
        Else
          If Versions(lpVersionIndex).Contents.Count = 0 Then
            Throw New InvalidOperationException("The version has no content element")
          End If
          ' This change will apply to the specified version of the document
          With Versions(lpVersionIndex).Contents(lpContentElementIndex)
            If String.Equals(.FileExtension, lpNewMimeType, StringComparison.InvariantCultureIgnoreCase) Then
              ' The proposed new mime type is the same as the current extension
              ' Do not make this change
              lblnReturnValue = False
            Else
              .MIMEType = lpNewMimeType
              lblnReturnValue = True
            End If
          End With
        End If

        ' Set up the return message
        ' For now the only reason for a failure is if the proposed new mime type 
        ' is the same as the extension.  If at some point we add other reasons 
        ' for failure we will need to take those into account for the message as well.

        If lblnReturnValue = True Then
          lpMessage = String.Format("Successfully changed mime type to '{0}' for document '{1}'.", lpNewMimeType, Me.ID)
          ApplicationLogging.WriteLogEntry(lpMessage, Reflection.MethodBase.GetCurrentMethod,
                                           TraceEventType.Information, 46723)
        Else
          lpMessage = String.Format("Declined to change mime type, the proposed mime type value '{0}', is the same as the current content file extension.", lpNewMimeType)
        End If

        Return lblnReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Changes the value of the specifed property in the document
    ''' </summary>
    ''' <param name="lpPropertyScope">Specifies whether the property to be changed is at the document level or the version level</param>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>If the property scope is set to the version level, all versions will be affected</remarks>
    Public Function ChangePropertyValue(ByVal lpPropertyScope As PropertyScope,
                                        ByVal lpName As String,
                                        ByVal lpNewValue As Object) As Boolean
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return ChangePropertyValue(lpPropertyScope, lpName, lpNewValue, Document.ALL_VERSIONS)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Changes the value of the specifed property in the document
    ''' </summary>
    ''' <param name="lpPropertyScope">Specifies whether the property to be changed is at the document level or the version level</param>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <param name="lpVersionIndex">The specific version on which to change the value</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>If the property scope is set to the version level, only the version specified by the lpVersionIndex will be affected.</remarks>
    Public Function ChangePropertyValue(ByVal lpPropertyScope As PropertyScope,
                                        ByVal lpName As String,
                                        ByVal lpNewValue As Object,
                                        ByVal lpVersionIndex As Integer) As Boolean

      'ApplicationLogging.WriteLogEntry("Enter Document::ChangePropertyValue", TraceEventType.Verbose)

      Try


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty
            If (StrComp(lpName, "DOCUMENT CLASS", CompareMethod.Text) = 0) OrElse (StrComp(lpName, "DOCUMENTCLASS", CompareMethod.Text) = 0) Then
              'If lpName.ToUpper = "DOCUMENT CLASS" Or lpName.ToUpper = "DOCUMENTCLASS" Then
              DocumentClass = lpNewValue
            Else
              Properties(lpName).ChangePropertyValue(lpNewValue)
            End If

          Case PropertyScope.VersionProperty
            Try
              If lpVersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
                ' This change will apply to all version of the document
                For Each lobjVersion As Version In Versions
                  lobjVersion.ChangePropertyValue(lpName, lpNewValue)
                Next

              Else
                ' This change will apply to the specified version of the document
                'Versions(lpVersionIndex).Properties(lpName).ChangePropertyValue(lpNewValue)
                Versions(lpVersionIndex).ChangePropertyValue(lpName, lpNewValue)

              End If
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              'ApplicationLogging.WriteLogEntry("Exit Document::ChangePropertyValue", TraceEventType.Verbose)
              Throw New InvalidOperationException("Could not change property value [" & ex.Message & "]", ex)
            End Try

        End Select

        'ApplicationLogging.WriteLogEntry("Exit Document::ChangePropertyValue", TraceEventType.Verbose)

        Return True
        ' Change the value of the specified property

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function ClearPropertyValue(ByVal lpPropertyScope As PropertyScope,
                                        ByVal lpName As String,
                                        ByVal lpVersionIndex As Integer) As Boolean
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty
            If (StrComp(lpName, "DOCUMENT CLASS", CompareMethod.Text) = 0) OrElse (StrComp(lpName, "DOCUMENTCLASS", CompareMethod.Text) = 0) Then
              'If lpName.ToUpper = "DOCUMENT CLASS" Or lpName.ToUpper = "DOCUMENTCLASS" Then
              DocumentClass = String.Empty
            Else
              Properties(lpName).ClearPropertyValue()
            End If

          Case PropertyScope.VersionProperty
            If lpVersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
              ' This change will apply to all version of the document
              For Each lobjVersion As Version In Versions
                lobjVersion.ClearPropertyValue(lpName)
              Next

            Else
              ' This change will apply to the specified version of the document
              'Versions(lpVersionIndex).Properties(lpName).ChangePropertyValue(lpNewValue)
              Versions(lpVersionIndex).ClearPropertyValue(lpName)
            End If
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a new Version object with a parent Document reference set to this document.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Does not add the new version to the Versions collection.</remarks>
    Public Function CreateVersion() As Version
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lobjVersion As New Version(Me)

        ' Versions.Add(lobjVersion)

        Return lobjVersion

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a property and adds it to all versions of the document.
    ''' </summary>
    Public Function CreateProperty(ByVal lpName As String, Optional ByVal lpValue As Object = Nothing,
                                   Optional ByVal lpValueType As PropertyType = PropertyType.ecmString,
                                   Optional ByVal lpPropertyScope As PropertyScope = PropertyScope.VersionProperty,
                                   Optional ByVal lpFirstVersionOnly As Boolean = False) As Boolean

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Method", TraceEventType.Verbose)
        If lpFirstVersionOnly = True Then
          Return CreateProperty(lpName, lpValue, lpValueType, lpPropertyScope, VersionScope.FirstVersionOnly)
        Else
          Return CreateProperty(lpName, lpValue, lpValueType, lpPropertyScope, VersionScope.AllVersions)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Creates a property and adds it to the specified versions of the document.
    ''' </summary>
    Public Function CreateProperty(ByVal lpName As String,
                                   ByVal lpValue As Object,
                                   ByVal lpValueType As PropertyType,
                                   ByVal lpPropertyScope As PropertyScope,
                                   ByVal lpVersionScope As VersionScope) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return CreateProperty(lpName, lpValue, lpValueType, lpPropertyScope, lpVersionScope, True)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a property and adds it to the specified versions of the document.
    ''' </summary>
    Public Function CreateProperty(ByVal lpName As String, ByVal lpValue As Object,
                                   ByVal lpValueType As PropertyType,
                                   ByVal lpPropertyScope As PropertyScope,
                                   ByVal lpVersionScope As VersionScope,
                                   ByVal lpPersistent As Boolean) As Boolean

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return CreateProperty(lpName,
                              lpValue,
                              lpValueType,
                              Cardinality.ecmSingleValued,
                              lpPropertyScope,
                              lpVersionScope,
                              lpPersistent)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::CreateProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Creates a property and adds it to the specified versions of the document.
    ''' </summary>
    Public Function CreateProperty(ByVal lpName As String,
                                   ByVal lpValue As Object,
                                   ByVal lpValueType As PropertyType,
                                   ByVal lpCardinality As Cardinality,
                                   ByVal lpPropertyScope As PropertyScope,
                                   ByVal lpVersionScope As VersionScope,
                                   ByVal lpPersistent As Boolean) As Boolean

      Try

        Dim lobjNewProperty As ECMProperty = Nothing

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Document::CreateProperty", TraceEventType.Verbose)

        ' Create the new property
        lobjNewProperty = PropertyFactory.Create(lpValueType, lpName, lpName, lpCardinality, lpValue, lpPersistent)

        Select Case lpPropertyScope
          Case PropertyScope.VersionProperty
            ' Create a new property for each version and optionally set its value.
            Select Case lpVersionScope
              Case VersionScope.FirstVersionOnly
                Versions(0).Properties.Add(lobjNewProperty)

              Case VersionScope.LastVersionOnly
                If Versions.Count > 0 Then
                  Versions(Versions.Count - 1).Properties.Add(lobjNewProperty)
                End If

              Case VersionScope.AllVersions
                For Each lobjVersion As Version In Versions
                  lobjNewProperty = PropertyFactory.Create(lpValueType, lpName, lpName, lpCardinality, lpValue, lpPersistent)
                  lobjVersion.Properties.Add(lobjNewProperty)
                Next

            End Select

            ApplicationLogging.WriteLogEntry("Exit Document::CreateProperty", TraceEventType.Verbose)
            Return True

          Case PropertyScope.DocumentProperty
            ' Create a new property for the document as a whole.
            Properties.Add(lobjNewProperty)

            ApplicationLogging.WriteLogEntry("Exit Document::CreateProperty", TraceEventType.Verbose)
            Return True

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::CreateProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Deletes all the content files for each version via the ContentPath (if they exist).
    ''' </summary>
    Public Sub DeleteContentPaths()
      Try
        If Me.Versions IsNot Nothing Then

          For Each lobjVersion As Version In Me.Versions
            lobjVersion.DeleteContentFiles()
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Deletes a property from all versions of the document.
    ''' </summary>
    Public Function DeleteProperty(ByVal lpPropertyScope As PropertyScope, ByVal lpName As String) As Boolean

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Document::DeleteProperty", TraceEventType.Verbose)

        Select Case lpPropertyScope
          Case PropertyScope.VersionProperty

            ' Delete the property for each version if it exists.
            For Each lobjVersion As Version In Versions
              If lobjVersion.Properties.PropertyExists(lpName) = False Then
                ' The property does not exist
                ApplicationLogging.WriteLogEntry(
  String.Format("Unable to delete property '{0}': The property does not exist in the current version.",
    lpName), TraceEventType.Warning, 2224)
                Return False
              Else
                lobjVersion.Properties.Delete(lpName)
              End If

            Next

          Case PropertyScope.DocumentProperty

            If PropertyExists(PropertyScope.DocumentProperty, lpName) Then
              Properties.Delete(lpName)
            Else
              ' The property does not exist
              ApplicationLogging.WriteLogEntry(
  String.Format("Unable to delete property '{0}': The property does not exist in the current document.",
    lpName), TraceEventType.Warning, 2224)
              Return False
            End If

          Case PropertyScope.BothDocumentAndVersionProperties, PropertyScope.AllProperties

            If PropertyExists(PropertyScope.DocumentProperty, lpName) Then
              Properties.Delete(lpName)
            End If

            ' Delete the property for each version if it exists.
            For Each lobjVersion As Version In Versions
              If lobjVersion.Properties.PropertyExists(lpName) Then
                lobjVersion.Properties.Delete(lpName)
              End If

            Next

        End Select


        ApplicationLogging.WriteLogEntry("Exit Document::DeleteProperty", TraceEventType.Verbose)

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::DeleteProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Renames a property in all versions of the document.
    ''' </summary>
    ''' <param name="lpPropertyScope">The scope of the property (Document, Version, etc.)</param>
    ''' <param name="lpCurrentName">The curent name of the property</param>
    ''' <param name="lpNewName">The new name for the property</param>
    ''' <returns>True if successful, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function RenameProperty(ByVal lpPropertyScope As PropertyScope, ByVal lpCurrentName As String, ByVal lpNewName As String) As Boolean

      Try

        Return RenameProperty(lpPropertyScope, lpCurrentName, lpNewName, String.Empty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::RenameProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Renames a property in all versions of the document.
    ''' </summary>
    ''' <param name="lpPropertyScope">The scope of the property (Document, Version, etc.)</param>
    ''' <param name="lpCurrentName">The curent name of the property</param>
    ''' <param name="lpNewName">The new name for the property</param>
    ''' <param name="lpErrorMessage">A string reference to capture any error messages</param>
    ''' <returns>True if successful, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function RenameProperty(ByVal lpPropertyScope As PropertyScope,
                               ByVal lpCurrentName As String,
                               ByVal lpNewName As String,
                               ByRef lpErrorMessage As String) As Boolean

      Try

        Dim lobjTargetProperty As ECMProperty = Nothing

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Document::RenameProperty", TraceEventType.Verbose)

        Select Case lpPropertyScope
          Case PropertyScope.VersionProperty
            For Each lobjVersion As Version In Versions
              If lobjVersion.Properties.PropertyExists(lpCurrentName) Then
                lobjTargetProperty = lobjVersion.Properties.ItemByName(lpCurrentName)
                lobjTargetProperty.Rename(lpNewName)
              Else
                ' The property does not exist
                lpErrorMessage = String.Format(
  "Unable to rename property '{0}' to '{1}': The property does not exist in the current version.",
    lpCurrentName, lpNewName)
                ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning, 2222)
                Return False
              End If
            Next

            ApplicationLogging.WriteLogEntry("Exit Document::RenameProperty", TraceEventType.Verbose)

            Return True

          Case PropertyScope.DocumentProperty
            If PropertyExists(PropertyScope.DocumentProperty, lpCurrentName) Then
              Properties(lpCurrentName).Name = lpNewName
            Else
              ' The property does not exist
              lpErrorMessage = String.Format(
    "Unable to rename property '{0}' to '{1}': The property does not exist in the current document.",
      lpCurrentName, lpNewName)
              ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning, 2223)
              Return False
            End If

            Return True

          Case PropertyScope.BothDocumentAndVersionProperties, PropertyScope.AllProperties

            For Each lobjVersion As Version In Versions
              If lobjVersion.Properties.PropertyExists(lpCurrentName) Then
                'lobjVersion.Properties(lpCurrentName).Name = lpNewName
                lobjTargetProperty = lobjVersion.Properties.ItemByName(lpCurrentName)
                lobjTargetProperty.Rename(lpNewName)
              End If
            Next

            If PropertyExists(lpCurrentName) Then
              'Properties(lpCurrentName).Name = lpNewName
              lobjTargetProperty = Properties.ItemByName(lpCurrentName)
              lobjTargetProperty.Rename(lpNewName)
            End If

        End Select

      Catch ex As Exception
        lpErrorMessage = ex.Message
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::RenameProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    Public Function TrimProperty(ByVal lpPropertyScope As PropertyScope,
                             ByVal lpPropertyName As String,
                             ByVal lpTrimType As TrimPropertyType,
                             Optional ByVal lpTrimCharacter As String = " ") As Boolean


#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

      Dim lstrPropertyValue As String
      ApplicationLogging.WriteLogEntry("Enter Document::TrimProperty", TraceEventType.Verbose)
      Try
        Select Case lpPropertyScope
          Case PropertyScope.VersionProperty
            For Each lobjVersion As Version In Versions
              lstrPropertyValue = lobjVersion.Properties(lpPropertyName).Value
              lstrPropertyValue = TrimValue(lstrPropertyValue, lpTrimType, lpTrimCharacter)
              lobjVersion.Properties(lpPropertyName).Value = lstrPropertyValue
            Next
            ApplicationLogging.WriteLogEntry("Exit Document::TrimProperty", TraceEventType.Verbose)
            Return True
          Case PropertyScope.DocumentProperty
            lstrPropertyValue = Properties(lpPropertyName).Value
            lstrPropertyValue = TrimValue(lstrPropertyValue, lpTrimType, lpTrimCharacter)
            Properties(lpPropertyName).Value = lstrPropertyValue
            Return True
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::TrimProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    Public Sub RemoveProperty(ByVal lpProperty As ECMProperty)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If PropertyExists(PropertyScope.DocumentProperty, lpProperty.Name) = True Then
          Me.Properties.Remove(lpProperty)
        End If
        If PropertyExists(PropertyScope.VersionProperty, lpProperty.Name) = True Then
          For Each lobjVersion As Version In Me.Versions
            If lobjVersion.Properties.PropertyExists(lpProperty.Name) Then
              lobjVersion.Properties.Remove(lpProperty)
            End If
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function AllVersionsHaveContent() As Boolean
      Try

        For Each lobjVersion As Version In Versions
          If lobjVersion.HasContent = False Then
            Return False
          End If
        Next

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Removes all versions without content and returns the number of versions removed.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RemoveVersionsWithoutContent() As Integer
      Try

        Dim lblnSuccess As Boolean
        Dim lintRemovalCount As Integer

        Do Until AllVersionsHaveContent()
          For Each lobjVersion As Version In Versions
            If lobjVersion.HasContent = False Then
              lblnSuccess = lobjVersion.RemoveFromDocument()
              If lblnSuccess = True Then
                lintRemovalCount += 1
                Exit For
              End If
            End If
          Next
        Loop

        ' Remove the versions without content in all of the related documents as well
        For Each lobjRelationship As Relationship In Relationships
          If lobjRelationship.RelatedDocument IsNot Nothing Then
            lintRemovalCount += lobjRelationship.RelatedDocument.RemoveVersionsWithoutContent
          End If
        Next

        If ContainsInvalidRelationships() Then
          Dim lintRemovedRelationshipsCount As Integer = RemoveInvalidRelationships()
        End If

        If lintRemovalCount = 1 Then
          ApplicationLogging.WriteLogEntry(String.Format("Removed {0} version without content from document '{1}'",
                                               lintRemovalCount, ID), TraceEventType.Information, 62301)
        ElseIf lintRemovalCount > 1 Then
          ApplicationLogging.WriteLogEntry(String.Format("Removed {0} versions without content from document '{1}'",
                                               lintRemovalCount, ID), TraceEventType.Information, 62301)
        End If

        Return lintRemovalCount

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Removes all invalid relationships and returns the 
    ''' number of relationships removed from the document.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RemoveInvalidRelationships() As Integer
      Try
        Dim lblnSuccess As Boolean
        Dim lintRemovalCount As Integer

        If ContainsInvalidRelationships() Then
          Do Until ContainsInvalidRelationships() = False
            For Each lobjRelationship As Relationship In Relationships
              If lobjRelationship.IsValid(Me) = False Then
                lblnSuccess = lobjRelationship.RemoveFromDocument(Me)
                If lblnSuccess = True Then
                  lintRemovalCount += 1
                  Exit For
                End If
              End If
            Next
          Loop
        Else
          Return 0
        End If

        If lintRemovalCount = 1 Then
          ApplicationLogging.WriteLogEntry(String.Format("Removed {0} invalid relationship from document '{1}'",
                                               lintRemovalCount, ID), TraceEventType.Information, 62302)
        ElseIf lintRemovalCount > 1 Then
          ApplicationLogging.WriteLogEntry(String.Format("Removed {0} invalid relationships from document '{1}'",
                                               lintRemovalCount, ID), TraceEventType.Information, 62302)
        End If

        Return lintRemovalCount

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    '' ''' <summary>
    '' ''' Returns the number of invalid relationships.
    '' ''' </summary>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public Function InvalidRelationshipCount() As Integer
    ''  Try
    ''    If Relationships Is Nothing OrElse Relationships.Count = 0 Then
    ''      Return 0
    ''    Else
    ''      Return Relationships.InvalidRelationshipCount(Me)
    ''    End If
    ''  Catch ex As Exception
    ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''    ' Re-throw the exception to the caller
    ''    Throw
    ''  End Try
    ''End Function

    ''' <summary>
    ''' Checks to see if there are any invalid relationships.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ContainsInvalidRelationships() As Boolean
      Try
        If Relationships Is Nothing OrElse Relationships.Count = 0 Then
          Return False
        Else
          Return Not Relationships.AllValid(Me)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Sets or changes a property value in all versions of the document.
    ''' </summary>
    Public Function SetPropertyValue(ByVal lpPropertyScope As PropertyScope,
                                 ByVal lpPropertyName As String, ByVal lpPropertyValue As Object) As Boolean


      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Document::SetPropertyValue", TraceEventType.Verbose)

        Return SetPropertyValue(lpPropertyScope, lpPropertyName, lpPropertyValue, False)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ApplicationLogging.WriteLogEntry("Exit Document::SetPropertyValue", TraceEventType.Verbose)
      End Try

    End Function

    Public Function SetPropertyValue(ByVal lpPropertyScope As PropertyScope,
                                 ByVal lpPropertyName As String,
                                 ByVal lpPropertyValue As Object, ByVal lpCreateProperty As Boolean,
                                 Optional ByVal lpPropertyType As PropertyType = PropertyType.ecmString) As Boolean

      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Enter Document::SetPropertyValue", TraceEventType.Verbose)

        SetPropertyValue(lpPropertyScope, lpPropertyName, lpPropertyValue, lpCreateProperty, lpPropertyType, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ApplicationLogging.WriteLogEntry("Exit Document::SetPropertyValue", TraceEventType.Verbose)
      End Try
    End Function

    Public Function SetPropertyValue(ByVal lpPropertyScope As PropertyScope,
                             ByVal lpPropertyName As String,
                             ByVal lpPropertyValue As Object, ByVal lpCreateProperty As Boolean,
                             ByVal lpPropertyType As PropertyType,
                             ByVal lpPersistent As Boolean) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Document::SetPropertyValue", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If PropertyExists(lpPropertyScope, lpPropertyName) = False Then
          ' The property does not currently exist
          If lpCreateProperty = True Then
            Dim lobjNewProperty As ECMProperty = PropertyFactory.Create(lpPropertyType, lpPropertyName, lpPropertyName, Cardinality.ecmSingleValued, lpPropertyValue)
            Select Case lpPropertyScope
              Case PropertyScope.DocumentProperty
                'Properties.Add(New ECMProperty(lpPropertyType, lpPropertyName, lpPropertyValue, lpPersistent))
                Properties.Add(lobjNewProperty)
                Return True
              Case PropertyScope.VersionProperty
                'GetLatestVersion.Properties.Add(New ECMProperty(lpPropertyType, lpPropertyName, lpPropertyValue, lpPersistent))
                GetLatestVersion.Properties.Add(lobjNewProperty)
                'GetLatestVersion.SetPropertyValue(lpPropertyName, lpPropertyValue, lpCreateProperty, lpPropertyType)
                Return True
            End Select
          End If
          Throw New InvalidOperationException(String.Format(
                                                  "Cannot set the value of '{0}' to '{1}' as the property does not exist in the document and the parameter lpCreateProperty was set to False.",
                                                  lpPropertyValue.ToString, lpPropertyName))
        Else

          Select Case lpPropertyScope

            Case PropertyScope.DocumentProperty
              Properties(lpPropertyName).Value = lpPropertyValue
              Return True

            Case PropertyScope.VersionProperty
              For Each lobjVersion As Version In Versions
                lobjVersion.Properties(lpPropertyName).Value = lpPropertyValue
              Next
              Return True

            Case Else
              Throw New NotImplementedException("Setting Content Metadata values is not yet supported, coming soon...")
          End Select

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ApplicationLogging.WriteLogEntry("Exit Document::SetPropertyValue", TraceEventType.Verbose)
      End Try
    End Function

    ' ''' <summary>
    ' ''' Determine if a property exists in the document.
    ' ''' </summary>
    ' ''' <param name="lpName">The name of the property to check</param>
    ' ''' <returns></returns>
    ' ''' <remarks>Checks in both document and version properties</remarks>
    'Public Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return PropertyExists(PropertyScope.BothDocumentAndVersionProperties, lpName, lpCaseSensitive)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ' ''' <summary>
    ' ''' Determine if a property exists in the document.
    ' ''' </summary>
    ' ''' <param name="lpName">The name of the property to check</param>
    ' ''' <returns></returns>
    ' ''' <remarks>Checks in both document and version properties</remarks>
    'Public Function PropertyExists(ByVal lpName As String) As Boolean
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return PropertyExists(PropertyScope.BothDocumentAndVersionProperties, lpName, True)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="lpPropertyScope"></param>
    ' ''' <param name="lpName"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function PropertyExists(ByVal lpPropertyScope As PropertyScope, ByVal lpName As String) As Boolean

    '  If IsDisposed Then
    '    Throw New ObjectDisposedException(Me.GetType.ToString)
    '  End If

    '  Return PropertyExists(lpPropertyScope, lpName, True)

    'End Function

    ' ''' <summary>
    ' ''' Determine if a property exists in the document.
    ' ''' </summary>
    ' ''' <param name="lpName">The name of the property to check</param>
    ' ''' <param name="lpPropertyScope">Specifies whether to check document scoped properties, version scoped properties or both</param>
    ' ''' <remarks></remarks>
    'Public Function PropertyExists(ByVal lpPropertyScope As PropertyScope, ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean

    '  ApplicationLogging.WriteLogEntry("Enter Document::PropertyExists", TraceEventType.Verbose)
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Dim lblnPropertyExists As Boolean

    '    Select Case lpPropertyScope
    '      Case PropertyScope.DocumentProperty
    '        lblnPropertyExists = Properties.PropertyExists(lpName, lpCaseSensitive)

    '      Case PropertyScope.VersionProperty
    '        For Each lobjVersion As Version In Versions
    '          lblnPropertyExists = lobjVersion.Properties.PropertyExists(lpName, lpCaseSensitive)
    '          If lblnPropertyExists = True Then
    '            Exit For
    '          End If
    '        Next

    '      Case PropertyScope.BothDocumentAndVersionProperties
    '        lblnPropertyExists = Properties.PropertyExists(lpName, lpCaseSensitive)
    '        If lblnPropertyExists = True Then
    '          ' If we already found the property then we can exit here
    '          Return lblnPropertyExists
    '        End If

    '        For Each lobjVersion As Version In Versions
    '          lblnPropertyExists = lobjVersion.Properties.PropertyExists(lpName, lpCaseSensitive)
    '          If lblnPropertyExists = True Then
    '            Exit For
    '          End If
    '        Next

    '    End Select

    '    Return lblnPropertyExists

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ApplicationLogging.WriteLogEntry("Exit Document::PropertyExists", TraceEventType.Verbose)
    '  Finally
    '    ApplicationLogging.WriteLogEntry("Exit Document::PropertyExists", TraceEventType.Verbose)
    '  End Try

    'End Function

    ''' <summary>
    ''' Deletes all properties from the document where no values are present
    ''' </summary>
    ''' <param name="lpErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks>Used to simplify the document and to shrink the list of properties down to the smallest available size</remarks>
    Public Function DeletePropertiesWithoutValues(Optional ByRef lpErrorMessage As String = "") As Document

      'Dim lobjProperty As ECMProperty

      ApplicationLogging.WriteLogEntry("Enter Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ' Go through all the document properties
        Properties = DeletePropertiesWithoutValues(Properties, lpErrorMessage)

        ' Go though all of the version properties
        For Each lobjVersion As Version In Versions
          lobjVersion.Properties = DeletePropertiesWithoutValues(lobjVersion.Properties, lpErrorMessage)
        Next

        ApplicationLogging.WriteLogEntry("Exit Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
        Return Me

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        ApplicationLogging.WriteLogEntry("Exit Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
        Return Nothing
      End Try

    End Function

    ''' <summary>
    ''' Deletes all properties from the document where no values are present
    ''' </summary>
    ''' <param name="lpErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks>Used to simplify the document and to shrink the list of properties down to the smallest available size</remarks>
    Private Shared Function DeletePropertiesWithoutValues(ByVal lpProperties As ECMProperties,
                                               Optional ByRef lpErrorMessage As String = "") As ECMProperties

      Dim lobjProperty As ECMProperty

      ApplicationLogging.WriteLogEntry("Enter Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
      Try

        For lintPropertyCounter As Integer = lpProperties.Count - 1 To 0 Step -1
          lobjProperty = lpProperties(lintPropertyCounter)

          If lobjProperty.Cardinality = Cardinality.ecmSingleValued Then
            If lobjProperty.Value Is Nothing Then
              lpProperties.Remove(lobjProperty)
            Else
              If lobjProperty.Value.ToString.Length = 0 Then
                lpProperties.Remove(lobjProperty)
              End If
            End If
          Else
            If lobjProperty.Values Is Nothing Then
              lpProperties.Remove(lobjProperty)
            End If
          End If
        Next

        ApplicationLogging.WriteLogEntry("Exit Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
        Return lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        ApplicationLogging.WriteLogEntry("lpErrorMessage", TraceEventType.Error)
        ApplicationLogging.WriteLogEntry("Exit Document::DeletePropertiesWithoutValues", TraceEventType.Verbose)
        Return Nothing
      End Try

    End Function

    ''' <summary>
    ''' Deletes all properties from the document with a Persistent value of False
    ''' </summary>
    ''' <param name="lpErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks>Used to clear out temporary properties that are about to go out 
    ''' of scope, such as in a transformation</remarks>
    Public Function DeleteTemporaryProperties(Optional ByRef lpErrorMessage As String = "") As Document

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ' Go through all the document properties
        Properties = DeleteTemporaryProperties(Properties, lpErrorMessage)

        ' Go though all of the version properties
        For Each lobjVersion As Version In Versions
          lobjVersion.Properties = DeleteTemporaryProperties(lobjVersion.Properties, lpErrorMessage)
        Next

        Return Me

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try

    End Function

    ''' <summary>
    ''' Deletes all properties from the document with a Persistent value of False
    ''' </summary>
    ''' <param name="lpErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks>Used to clear out temporary properties that are about to go out 
    ''' of scope, such as in a transformation</remarks>
    Public Shared Function DeleteTemporaryProperties(ByVal lpProperties As ECMProperties,
                                               Optional ByRef lpErrorMessage As String = "") As ECMProperties

      Dim lobjProperty As ECMProperty

      Try
        For lintPropertyCounter As Integer = lpProperties.Count - 1 To 0 Step -1
          lobjProperty = lpProperties(lintPropertyCounter)

          If lobjProperty.Persistent = False Then
            lpProperties.Remove(lobjProperty)
          End If
        Next

        Return lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Document Content Methods"

    ''' <summary>
    ''' Determines whether or not all the content specified in the document exists
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AllContentExists() As Boolean

      ' Does the content as specified for each version exist?
      ApplicationLogging.WriteLogEntry("Enter Document::AllContentExists", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lblnExists As Boolean
        For Each lobjVersion As Version In Versions
          For Each lobjContent As Content In lobjVersion.Contents
            Select Case lobjContent.StorageType
              Case Content.StorageTypeEnum.Reference
                lblnExists = IO.File.Exists(lobjContent.ContentPath)
                If lblnExists = False Then
                  Return False
                End If
              Case Content.StorageTypeEnum.EncodedUnCompressed, Content.StorageTypeEnum.EncodedCompressed
                If lobjContent.Data Is Nothing Or lobjContent.Data.Length = 0 Then
                  '  We have no data for this content
                  Return False
                End If
            End Select
          Next
        Next
        ApplicationLogging.WriteLogEntry("Exit Document::AllContentExists", TraceEventType.Verbose)
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::AllContentExists", TraceEventType.Verbose)
        Return False
      End Try
    End Function

    ''' <summary>
    ''' Determines whether or not the document has any annotations.
    ''' </summary>
    ''' <returns>True if annotations are present, otherwise false.</returns>
    ''' <remarks>If a document has no content, there will be no annotations.</remarks>
    Public Function HasAnnotations() As Boolean
      Try

        If HasContent() = False Then
          Return False
        Else
          ' See if any of the content has annotations

          ' Look through each version
          For Each lobjVersion As Version In Versions
            ' Look through each content
            For Each lobjContent As Content In lobjVersion.Contents
              ' Check to see if the annotation count is non zero
              If lobjContent.Annotations.Count > 0 Then
                ' If there is at least one annotation then the document has annotations.
                ' We do not need to look any further.
                Return True
              End If
            Next
          Next

          ' If we made it this far then we did not find any annotations.
          Return False

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a copy of the document with all but the latest version removed.  
    ''' The contents of the latest version are replaced with a cpf of the original document.   
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Layer() As Document
      Try
        Return Layer(False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a copy of the document with all but the latest version removed.  
    ''' The contents of the latest version are replaced with a cpf of the original document.
    ''' </summary>
    ''' <param name="lpPackageAsJson">Specifies whether or not to package the document using json.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Layer(lpPackageAsJson As Boolean) As Document
      Try
        Dim lobjPackagedDocument As Document = Me.Clone(True)
        Select Case lobjPackagedDocument.Versions.Count
          Case 0
            lobjPackagedDocument.Versions.Add(New Version(lobjPackagedDocument))
          Case Is > 1
            ' Remove all but the latest version
            lobjPackagedDocument.Versions.Reverse()
            For lintVersionCounter As Integer = lobjPackagedDocument.Versions.Count - 1 To 1 Step -1
              lobjPackagedDocument.Versions.RemoveAt(lintVersionCounter)
            Next
        End Select
        If lpPackageAsJson Then
          lobjPackagedDocument.LatestVersion.Contents.Add(Me.ToJsonArchiveStream(),
  String.Format("{0}.jcpf", Helper.CleanFile(Me.ID, "`")), Content.StorageTypeEnum.Reference)
        Else
          lobjPackagedDocument.LatestVersion.Contents.Add(Me.ToArchiveStream,
  String.Format("{0}.cpf", Helper.CleanFile(Me.ID, "`")), Content.StorageTypeEnum.Reference)
        End If


        Return lobjPackagedDocument

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a copy of the specified document with all but the latest version removed.  
    ''' The contents of the latest version are replaced with a cpf of the original document.
    ''' </summary>
    ''' <param name="document">The document to layer.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Layer(document As Document) As Document
      Try
        Return document.Layer()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a copy of the specified document with all but the latest version removed.  
    ''' The contents of the latest version are replaced with a cpf of the original document.    ''' </summary>
    ''' <param name="document">The document to layer.</param>
    ''' <param name="packageAsJson">Specifies whether or not to package the document using json.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Layer(document As Document, packageAsJson As Boolean) As Document
      Try
        Return document.Layer()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetLargestFileSize() As FileSize

      ApplicationLogging.WriteLogEntry("Enter Document::GetLargestFileSize", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lobjLargestVersionFileSize As FileSize
        Dim lobjLargestFileSize As New FileSize(0)

        For Each lobjVersion As Version In Versions
          lobjLargestVersionFileSize = lobjVersion.GetLargestFileSize
          If lobjLargestVersionFileSize.CompareTo(lobjLargestFileSize) > 0 Then
            lobjLargestFileSize = lobjLargestVersionFileSize
          End If
        Next
        ApplicationLogging.WriteLogEntry("Exit Document::GetLargestFileSize", TraceEventType.Verbose)

        Return lobjLargestFileSize

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::GetLargestFileSize", TraceEventType.Verbose)
        Return Nothing
      End Try

    End Function

    '' ''' <summary>
    '' ''' Sets the current StorageType value and encodes all the content data per the specified storage type
    '' ''' </summary>
    '' ''' <param name="lpStorageType">The storage type for the document content</param>
    '' ''' <remarks></remarks>
    ''Public Sub EncodeAllContents(ByVal lpStorageType As Content.StorageTypeEnum)
    ''  Try

    ''    If IsDisposed Then
    ''      Throw New ObjectDisposedException(Me.GetType.ToString)
    ''    End If

    ''    '  Only proceed if the current storage type does not match the one specified here
    ''    If StorageType = lpStorageType Then
    ''      Exit Sub
    ''    End If

    ''    '  First set the document storage type to match the one specified here
    ''    StorageType = lpStorageType

    ''    '  Set the data accordingly for each content in every version
    ''    For Each lobjVersion As Version In Me.Versions
    ''      For Each lobjContent As Content In lobjVersion.Contents
    ''        lobjContent.SetData(lpStorageType)
    ''      Next
    ''    Next

    ''  Catch ex As Exception
    ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''    ' Re-throw the exception to the caller
    ''    Throw
    ''  End Try
    ''End Sub

    '' ''' <summary>
    '' ''' Encodes all the content data per for current storage type
    '' ''' </summary>
    '' ''' <remarks></remarks>
    ''Public Sub EncodeAllContents()

    ''  If IsDisposed Then
    ''    Throw New ObjectDisposedException(Me.GetType.ToString)
    ''  End If

    ''  EncodeAllContents(Me.StorageType)

    ''End Sub

    Private Sub UpdateContentLocations(ByVal lpContentReferenceBehavior As ContentReferenceBehavior)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Select Case lpContentReferenceBehavior

          Case ContentReferenceBehavior.Move, ContentReferenceBehavior.Copy
            If Me.SerializationPath.Length > 0 Then
              For Each lobjVersion As Version In Me.Versions
                For Each lobjContent As Content In lobjVersion.Contents
                  lobjContent.UpdateContentLocation(lpContentReferenceBehavior)
                Next
              Next
            End If

          Case Else
            ' Do nothing
            Exit Sub
        End Select


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

    ''' <summary>
    ''' Updates the MIMEType of all the content in all the versions 
    ''' based on the registered extensions in the Windows Registry.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateAllContentMimeTypes()
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        UpdateAllContentMimeTypes(String.Empty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Updates the MIMEType of all the content in all the versions 
    ''' based on the registered extensions in the Windows Registry.
    ''' </summary>
    ''' <param name="lpNewMimeType">The new MIMEType to assign.</param>
    ''' <remarks>
    ''' If no new MIMEType is specified, the mime type will be 
    ''' resolved to the registered extensions in the Windows registry.
    ''' </remarks>
    Public Sub UpdateAllContentMimeTypes(ByVal lpNewMimeType As String)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        For Each lobjVersion As Version In Me.Versions
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContent.UpdateMimeType(lpNewMimeType)
          Next
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Iterates through each ContentItem and sets's it's 
    ''' ContentPath relative to the parent document's SerializationPath
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetContentPathsFromSerializationPath()
      Try

        If Me.SerializationPath.Length = 0 Then
          'Throw New InvalidOperationException("Unable to set content paths, the SerializationPath property is not set.")
          ApplicationLogging.WriteLogEntry(String.Format("Unable to set content paths for document '{0}', the SerializationPath property is not set.", Me.ID), TraceEventType.Warning, 61273)
          Exit Sub
        End If

        For Each lobjVersion As Version In Me.Versions
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContent.SetContentPathFromDocumentSerializationPath()
          Next
        Next


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    '    Public Sub DecodeAllContents(ByVal lpDecompress As Boolean)
    '      Try

    '        '  This only applies if the storage type is not currently by reference
    '        If Me.StorageType = Content.StorageTypeEnum.Reference Then
    '          Throw New InvalidOperationException("The current document storage type is by reference, there is nothing to decode")
    '        End If

    '        For Each lobjVersion As Version In Me.Versions
    '          For Each lobjContent As Content In lobjVersion.Contents
    '            If lobjContent.ContentPath IsNot Nothing AndAlso lobjContent.Data IsNot Nothing Then
    '              '  Decode all the content
    'lobjContent.
    '              Content.DecodeFileFromString(lobjContent.Data, lobjContent.ContentPath, lpDecompress)
    '            ElseIf lobjContent.ContentPath Is Nothing OrElse lobjContent.ContentPath.Length = 0 Then
    '              Dim lobjInvalidPathException As New InvalidPathException
    '              Throw New DocumentException(Me, lobjVersion.ID, "Unable to DecodeAllContents, one or more ContentPath values is invalid.", lobjInvalidPathException)
    '            End If
    '          Next
    '        Next
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

#End Region

#Region "Document Transformation Methods"

    Public Function Transform(ByVal lpTransformation As Transformation, Optional ByRef lpErrorMessage As String = "") As Document

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return lpTransformation.TransformDocument(Me, lpErrorMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Document Serialization Methods"

    Public Function SetSerializationPath(ByVal lpSerializationPath As String) As String

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        mstrSerializationPath = lpSerializationPath
        Return SerializationPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ' Moved to an extension method in Ecmg.Cts.Validation
    'Public Function Validate(ByVal lpValidation As Validation) As ValidationResult

    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    Return lpValidation.ValidateDocument(Me)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Validates a document using the specified validation test
    '''' </summary>
    '''' <param name="lpValidationTest">The validation test to be run for the document.</param>
    '''' <returns></returns>
    '''' <remarks>Uses a live ContentSource connection for validation against the destination repository.</remarks>
    'Public Function Validate(ByVal lpValidationTest As ValidationTest) As ValidationResult
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    ' Configure the validation
    '    Dim lobjValidation As New Validation

    '    With lobjValidation
    '      ' Set the destination content source
    '      .Importer = Me.GetContentSource
    '      ' Add the test
    '      lobjValidation.Tests.Add(lpValidationTest)
    '      Return .ValidateDocument(Me)
    '    End With

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    ''''  Validates a document using the specified validation test
    '''' </summary>
    '''' <param name="lpValidationTest">The validation test to be run for the document.</param>
    '''' <param name="lpDestinationRepository">A repository object representing the destination repository to be tested against.</param>
    '''' <returns></returns>
    '''' <remarks>Uses an offline repository object reference for validation against the destination repository.</remarks>
    'Public Function Validate(ByVal lpValidationTest As ValidationTest, ByVal lpDestinationRepository As Repository) As ValidationResult
    '  Try

    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    ' Configure the validation
    '    Dim lobjValidation As New Validation(lpDestinationRepository)

    '    With lobjValidation
    '      ' Set the destination content source
    '      '.Importer = Me.ContentSource
    '      ' Add the test
    '      lobjValidation.Tests.Add(lpValidationTest)
    '      Return .ValidateDocument(Me)
    '    End With

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function


    ''' <summary>
    ''' Generates a path extension to be appended to the end of an ExportPath.
    ''' This allows large export jobs to be spread across multiple subfolders 
    ''' by saving to the extended path.
    ''' </summary>
    ''' <param name="lpWidth">The number of folders to be used</param>
    ''' <param name="lpDepth">The number of folder levels</param>
    ''' <returns>A string representing subfolders to be added to the current export path (i.e. "\8\12\2\009543825"</returns>
    ''' <remarks></remarks>
    Public Function PathExtension(ByVal lpWidth As Byte,
                              ByVal lpDepth As Byte) As String
      Try

        Dim lobjFolderPaths As IList(Of String) = Me.FolderPaths

        If lobjFolderPaths.Count > 0 Then
          Return GetPathExtension(Me.ID,
                        String.Format("{0}\{1}", lobjFolderPaths(0), Me.Name),
                        lpWidth, lpDepth)
        Else
          Return GetPathExtension(Me.ID, Me.Name, lpWidth, lpDepth)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Generates a path extension to be appended to the end of an ExportPath.
    ''' This allows large export jobs to be spread across multiple subfolders 
    ''' by saving to the extended path.
    ''' </summary>
    ''' <param name="documentId">The document identifier</param>
    ''' <param name="originalPath">The primary original repository 
    ''' folder path for the document including the document title, 
    ''' if the document is unfiled the document title could be used instead.</param>
    ''' <param name="width">The number of folders to be used</param>
    ''' <param name="depth">The number of folder levels</param>
    ''' <returns>A string representing subfolders to be added to the current export path (i.e. "\8\12\2\009543825"</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentOutOfRangeException">The specified depth must be less than the length provided for originalPath</exception>
    Public Shared Function GetPathExtension(ByVal documentId As String,
                                        ByVal originalPath As String,
                                        ByVal width As Byte,
                                        ByVal depth As Byte) As String
      Try

        Dim lobjExtensionBuilder As New StringBuilder()
        Dim lintFolderNumber As Integer
        Dim lintHash As Integer

        If depth > originalPath.Length Then
          Throw New ArgumentOutOfRangeException(NameOf(depth), depth,
  String.Format("The folder depth of '{0}' can not be greater than the string length '{1}' of the originalPath '{2}'", depth, originalPath.Length, originalPath))
        End If

        ' Loop through each depth level
        For lintDepthCounter As Byte = 1 To depth

          ' Compute a simple numeric hash for the original path
          ' On each pass through we will remove one letter from the end
          ' to get a new result
          lintHash = Encryption.SimpleHash.ComputeBasicNumber(
  originalPath.Substring(0, originalPath.Length - lintDepthCounter - 1))

          ' Figure out which folder we will use in this depth level
          lintFolderNumber = Math.Abs(lintHash Mod width)

          ' Add this folder to the level
          lobjExtensionBuilder.AppendFormat("\{0}", lintFolderNumber + 1)

        Next

        ' After we have completed all the levels add the document id as the last level
        lobjExtensionBuilder.AppendFormat("\{0}", documentId)

        ' Return the result
        Return lobjExtensionBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Basic Content Services Methods"

    ''' <summary>
    ''' Adds the document to the repository
    ''' </summary>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implements the "IBasicContentServicesProvider" interface.</remarks>
    Public Function Add() As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.AddDocument(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Adds the document to the repository
    ''' </summary>
    ''' <param name="lpFolderPath">The path to file the document into</param>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implements the "IBasicContentServicesProvider" interface.</remarks>
    Public Function Add(ByVal lpFolderPath As String) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.AddDocument(Me, lpFolderPath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks out the document from the repository
    ''' </summary>
    ''' <param name="lpDestinationFolder">The local or network folder to check the document out to</param>
    ''' <param name="lpOutputFileNames">A string array to capture the path(s) of the content element(s) checked out from the repository.</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>The destination folder should exist before calling this method.
    ''' Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implements the "IBasicContentServicesProvider" interface.</remarks>
    Public Function CheckOut(ByVal lpDestinationFolder As String, ByRef lpOutputFileNames As String()) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.CheckoutDocument(Me.ID, lpDestinationFolder, lpOutputFileNames)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualifed path to the file to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpContentPath As String, ByVal lpAsMajorVersion As Boolean) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.CheckinDocument(Me.ID, lpContentPath, lpAsMajorVersion)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpContentPaths">An array of strings containing the fully qualified paths of all of the content elements to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpContentPaths As String(),
                         ByVal lpAsMajorVersion As Boolean) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.CheckinDocument(Me.ID, lpContentPaths, lpAsMajorVersion)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Cancels the checkout of a document
    ''' </summary>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Public Function CancelCheckout() As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.CancelCheckoutDocument(Me.ID)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Copies out the content of the latest document to the specified destination folder
    ''' </summary>
    ''' <param name="lpDestinationFolder">The fully qualified path to the folder to copy the content file(s) to</param>
    ''' <param name="lpOutputFileNames">A string array returned by reference with the fully qualified file names of the content elements copied out</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>The destination folder should exist before calling this method.
    ''' Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implements the "IBasicContentServicesProvider" interface.</remarks>
    Public Function CopyOut(ByVal lpDestinationFolder As String, ByRef lpOutputFileNames As String()) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.CopyOutDocument(Me.ID, lpDestinationFolder, lpOutputFileNames)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub ClearAllContents()
      Try
        For Each lobjVersion As Version In Me.Versions
          If lobjVersion.HasContent Then
            lobjVersion.Contents.Clear()
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Declares the document as a record
    ''' </summary>
    ''' <param name="lpArgs">An Ecmg.Cts.Records.DeclareRecordArgs object reference.</param>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implements the "IRecordsManager" interface.</remarks>
    Public Overridable Function DeclareAsRecord(ByVal lpArgs As Object) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetRMProvider.DeclareRecord(lpArgs)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    ''' <summary>
    ''' Files the document in the specified path
    ''' </summary>
    ''' <param name="lpFolderPath">The path to file the document into</param>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implement the "IBasicContentServicesProvider" interface.</remarks>
    Public Function File(ByVal lpFolderPath As String) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If Me.ID.Length > 0 Then
          Return GetBcsProvider.FileDocument(Me.ID, lpFolderPath)
        ElseIf Me.ObjectID.Length > 0 Then
          Return GetBcsProvider.FileDocument(Me.ObjectID, lpFolderPath)
        Else
          Throw New DocumentException(Me, "Unable to file document, no document identifer is available.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Unfiles the document in the specified path
    ''' </summary>
    ''' <param name="lpFolderPath">The path to unfile the document from</param>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implement the "IBasicContentServicesProvider" interface.</remarks>
    Public Function Unfile(ByVal lpFolderPath As String) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If Me.ID.Length > 0 Then
          Return GetBcsProvider.UnFileDocument(Me.ID, lpFolderPath)
        ElseIf Me.ObjectID.Length > 0 Then
          Return GetBcsProvider.UnFileDocument(Me.ObjectID, lpFolderPath)
        Else
          Throw New DocumentException(Me, "Unable to unfile document, no document identifer is available.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Updates the document properties
    ''' </summary>
    ''' <param name="Args">The properties to update</param>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implement the "IBasicContentServicesProvider" interface.</remarks>
    Public Function UpdateProperties(ByVal Args As DocumentPropertyArgs) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.UpdateDocumentProperties(Args)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Deletes the document from the repository
    ''' </summary>
    ''' <returns>True if successful or False if not</returns>
    ''' <remarks>Requires a valid ContentSource.  The ContentSource property of the document must be set to 
    ''' a valid ContentSource object reference which must also implement the "IBasicContentServicesProvider" interface.</remarks>
    Public Function Delete() As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetBcsProvider.DeleteDocument(Me.ID)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ITableable Implementation"

    ''' <summary>
    ''' Creates a one-row DataTable with all of the document properties and the properties of the first version.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToDataTable() As DataTable Implements ITableable.ToDataTable
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return ToDataTable(0)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a one-row DataTable with all of the document properties and the properties of the specified version.
    ''' </summary>
    ''' <param name="lpVersionIndex">The version to output to the DataTable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToDataTable(ByVal lpVersionIndex As Integer) As DataTable
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If mobjDataTable IsNot Nothing Then
          Return mobjDataTable
        End If

        ' If there were no results then we will return a null reference
        'If Results.Count = 0 Then
        '  Return Nothing
        'End If

        ' Build the data table

        mobjDataTable = New DataTable("Document")

        mobjDataTable.Columns.Add("Id", GetType(String))
        CreateDataTableColumns(Me.Properties)
        CreateDataTableColumns(Me.Versions(lpVersionIndex).Properties)

        Dim lobjDataRow As DataRow

        ' Add the rows
        lobjDataRow = mobjDataTable.NewRow

        lobjDataRow.Item("Id") = Me.ID
        SetDataTableRow(Me.Properties, lobjDataRow)
        SetDataTableRow(Me.Versions(lpVersionIndex).Properties, lobjDataRow)


        mobjDataTable.Rows.Add(lobjDataRow)


        Return mobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates the DataRow for the document DataTable
    ''' </summary>
    ''' <param name="lpProperties">The properties collection of either the document or the version</param>
    ''' <param name="lobjDataRow"></param>
    ''' <remarks></remarks>
    Private Shared Sub SetDataTableRow(ByVal lpProperties As ECMProperties, ByVal lobjDataRow As DataRow)

      Try
        Dim s As String = ""
        For Each lobjProperty As ECMProperty In lpProperties

          Select Case lobjProperty.Cardinality

            Case Cardinality.ecmSingleValued

              If lobjProperty.HasValue Then
                lobjDataRow.Item(lobjProperty.SystemName) = lobjProperty.Value
              Else
                lobjDataRow.Item(lobjProperty.SystemName) = DBNull.Value
              End If

            Case Cardinality.ecmMultiValued
              s = String.Empty
              For Each lobjValue As Object In lobjProperty.Values
                If TypeOf (lobjValue) Is Value Then
                  s &= lobjValue.Value.ToString & vbCrLf
                Else
                  s &= lobjValue.ToString & vbCrLf
                End If
              Next
              'Added Length check by RKS 10/27/2008
              If (s.Length > 0) Then
                s = s.TrimEnd(vbCrLf)
              End If
              lobjDataRow.Item(lobjProperty.SystemName) = s
          End Select

        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Creates all the DataColumns for the document DataTable
    ''' </summary>
    ''' <param name="lpProperties">The properties collection of either the document or the version</param>
    ''' <remarks></remarks>
    Private Sub CreateDataTableColumns(ByVal lpProperties As ECMProperties)

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        With mobjDataTable

          ' Add our columns
          For Each lobjDataItem As ECMProperty In lpProperties

            'If the dataitem is already in the table, do not add again. RKS - 7/23/2012
            If mobjDataTable.Columns.Contains(lobjDataItem.SystemName) Then
              Continue For
            End If

            Select Case lobjDataItem.Type
              Case PropertyType.ecmString, PropertyType.ecmGuid
                .Columns.Add(lobjDataItem.SystemName, GetType(String))

              Case PropertyType.ecmDate
                .Columns.Add(lobjDataItem.SystemName, GetType(DateTime))

              Case PropertyType.ecmLong
                .Columns.Add(lobjDataItem.SystemName, GetType(Integer))

              Case PropertyType.ecmDouble
                .Columns.Add(lobjDataItem.SystemName, GetType(Double))

              Case PropertyType.ecmBoolean
                .Columns.Add(lobjDataItem.SystemName, GetType(Boolean))

                'Case PropertyType.ecmGuid
                '  .Columns.Add(lobjDataItem.Name, GetType(Guid))

              Case PropertyType.ecmObject
                .Columns.Add(lobjDataItem.SystemName, GetType(Object))

              Case PropertyType.ecmBinary
                .Columns.Add(lobjDataItem.SystemName, GetType(Object))

            End Select
          Next

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Overloads"

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

#End Region

#Region "Private Methods"

    'Private Sub InitializeObjectID()
    '  Try
    '    If PropertyExists("ObjectID") = False Then
    '      Properties.Add(New ECMProperty(PropertyType.ecmString, _
    '              "ObjectID", ""))
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ''' <summary>
    ''' Used to clean up all existing files related to this document from disc.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DeleteAllRelatedFiles(ByVal lpDocumentRootPath As String)
      Try

        Dim lstrVersionFolderPath As String = String.Empty
        Dim lobjSerializationFolder As DirectoryInfo = Nothing
        Dim lobjVersionFolders() As DirectoryInfo = Nothing

        If SerializationPath IsNot Nothing AndAlso
  SerializationPath.Length > 0 AndAlso
  Directory.Exists(Path.GetDirectoryName(SerializationPath)) Then

          ' If we did not serialize to a subfolder specific to the document then skip the operation.
          If Path.GetDirectoryName(SerializationPath).Contains(Me.ID) = False Then
            Exit Sub
          End If

          lobjSerializationFolder = New DirectoryInfo(Path.GetDirectoryName(SerializationPath))
          lobjVersionFolders = lobjSerializationFolder.GetDirectories()

        ElseIf lpDocumentRootPath IsNot Nothing AndAlso
  lpDocumentRootPath.Length > 0 AndAlso
  Directory.Exists(Path.GetDirectoryName(lpDocumentRootPath)) Then

          lobjSerializationFolder = New DirectoryInfo(lpDocumentRootPath)
          lobjVersionFolders = lobjSerializationFolder.GetDirectories()

        End If

        For lintVersionCounter As Integer = 0 To Versions.Count - 1
          For Each lobjContent As Content In Versions(lintVersionCounter).Contents
            lobjContent.DeleteContentFiles()
            If lobjVersionFolders IsNot Nothing AndAlso lobjVersionFolders.Length > 0 AndAlso lobjVersionFolders.Length > lintVersionCounter Then
              For Each lobjFileInfo As FileInfo In lobjVersionFolders(lintVersionCounter).GetFiles
                If String.Equals(lobjFileInfo.Name, lobjContent.FileName, StringComparison.InvariantCultureIgnoreCase) Then
                  lobjFileInfo.Delete()
                End If
              Next
            End If
          Next

        Next
        'For Each lobjVersion As Version In Versions
        '  For Each lobjContent As Content In lobjVersion.Contents
        '    lobjContent.DeleteContentFiles()
        '  Next
        'Next

        ' Try to clean up the version folders
        If lobjVersionFolders IsNot Nothing Then
          For Each lobjVersionFolder As DirectoryInfo In lobjVersionFolders
            Try
              If lobjVersionFolder.Exists AndAlso
  lobjVersionFolder.GetFiles.Length = 0 AndAlso
  lobjVersionFolder.GetDirectories.Length = 0 Then
                lobjVersionFolder.Delete()
              End If
            Catch UnauthorizedAccessEx As UnauthorizedAccessException
              ' Skip it and move on
              Continue For
            Catch ex As Exception
              ' Log it to the application log
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              ' Skip it and move on
              Continue For
            End Try
          Next
        End If

        ' Try to delete the cdf file if it exists
        If SerializationPath IsNot Nothing AndAlso IO.File.Exists(SerializationPath) Then
          IO.File.Delete(SerializationPath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetContentCount() As Integer
      Try
        Dim lintContentCount As Integer

        For Each lobjRelationship As Relationship In Relationships
          If lobjRelationship.RelatedDocument IsNot Nothing Then
            lintContentCount += lobjRelationship.RelatedDocument.ContentCount
          End If
        Next

        For Each lobjVersion As Version In Versions
          lintContentCount += lobjVersion.ContentCount
        Next

        Return lintContentCount

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub GetRelatedDocuments(lpDocument As Document, lpZipFile As ZipFile)
      Try

        Dim lstrRelatedDocumentPackageName As String = String.Empty

        ' First try to extract the contents from any related documents
        For Each lobjRelationship As Relationship In lpDocument.Relationships
          If lobjRelationship.RelatedDocument IsNot Nothing Then
            ' We need to extract the related document cpf from the parent cpf

            Dim lobjRelatedDocument As Document = Nothing

            If RelatedDocuments.Contains(lobjRelationship.RelatedDocument.ID, lobjRelatedDocument) = False Then

              ' Get the cpf name
              lstrRelatedDocumentPackageName = String.Format("Relationships/{0}.cpf", lobjRelationship.RelatedDocument.ID)

              If lpZipFile.ContainsEntry(lstrRelatedDocumentPackageName) Then

                ' Get the cpf entry from the parent cpf
                Dim lobjRelatedZipEntry As ZipEntry = lpZipFile.Item(lstrRelatedDocumentPackageName)
                ' Open the child cpf into a stream
                Dim lobjRelatedDocumentStream As New MemoryStream
                lobjRelatedZipEntry.Extract(lobjRelatedDocumentStream)
                lobjRelatedDocumentStream.Position = 0
                ' Create the child zip file from the stream
                Dim lobjRelatedZipFile As ZipFile = ZipFile.Read(lobjRelatedDocumentStream)

                lobjRelatedDocument = New Document(lobjRelatedDocumentStream)

                RelatedDocuments.Add(lobjRelatedDocument)

              Else
                ' Make a notation in the log and move on.
                ' We may need to change this later but for now we do 
                ' not want to hold up the processing of this document
                Dim lstrErrorMessage As String = String.Format("Unable to extract content for related document '{0}' using entry '{1}'.  The entry could not be found in parent archive '{2}.cpf'.",
                lobjRelationship.RelatedDocument.Name, lstrRelatedDocumentPackageName, lpDocument.ID)
                Throw New Exceptions.PropertyDoesNotExistException(lstrErrorMessage, lstrRelatedDocumentPackageName)
              End If

            End If

            lobjRelationship.RelatedDocument = lobjRelatedDocument

          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the Document object in a CDF file.
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified output path</param>
    ''' <param name="lpContentReferenceBehavior">
    ''' Specifies whether or not the content files will 
    ''' be moved or copied under the new serialization path.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub Save(ByRef lpFilePath As String,
             Optional ByVal lpContentReferenceBehavior As ContentReferenceBehavior = ContentReferenceBehavior.Copy)
      Try

        Save(lpFilePath, "", False, "", lpContentReferenceBehavior)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the Document object in a CDF file.
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified output path</param>
    ''' <param name="lpFileExtension">Used if desired to override the default file extension of .cdf</param>
    ''' <param name="lpWriteProcessingInstruction">Specifies whether or not an XML write procesing instruction is to be provided</param>
    ''' <param name="lpStyleSheetPath">If the value of lpWriteProcessingInstruction is set to true, then the specified style sheet path is inserted into the file</param>
    ''' <param name="lpContentReferenceBehavior">Specifies whether or not the cont files will be moved or copied under the new serialization path</param>
    ''' <remarks></remarks>
    Private Sub Save(ByRef lpFilePath As String,
                 ByVal lpFileExtension As String,
                 ByVal lpWriteProcessingInstruction As Boolean,
                 ByVal lpStyleSheetPath As String,
                 Optional ByVal lpContentReferenceBehavior As ContentReferenceBehavior = ContentReferenceBehavior.Move)
      Try

        mstrCurrentPath = lpFilePath

        ' Set the extension if necessary
        ' We want the default file extension on an ecmdocument to be .cdf for
        ' "Content Definition File"

        If lpFileExtension.Length = 0 Then
          ' No override was provided
          If lpFilePath.EndsWith(DefaultFileExtension) = False Then
            lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DefaultFileExtension
          End If

        End If

        SetDocumentReferences()

        ' Called to provide an audit trail for subsequent re-serialization after transformation
        SetSerializationPath(lpFilePath)

        '  Encode the contents as appropriate for the storage type
        EncodeAllContents()

        '  Update the content locations before we write out the cdf file.
        If StorageType = Content.StorageTypeEnum.Reference AndAlso
 SerializationPath IsNot Nothing AndAlso
 SerializationPath <> String.Empty Then

          UpdateContentLocations(lpContentReferenceBehavior)

        End If

        UpdateHeader()

        If IO.Path.GetExtension(lpFilePath).Replace(".", String.Empty) = JSON_CONTENT_DEFINITION_FILE_EXTENSION Then
          Me.ToJsonFile(lpFilePath)
        Else
          If lpWriteProcessingInstruction = True Then
            Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
          Else
            ' Pass along the culture
            If Not String.IsNullOrEmpty(OriginalLocale) Then
              Serializer.Serialize.XmlFile(Me, lpFilePath, , , , , , , New CultureInfo(OriginalLocale))
            Else
              Serializer.Serialize.XmlFile(Me, lpFilePath)
            End If

          End If
        End If

        If StorageType <> Content.StorageTypeEnum.Reference Then
          '  Delete the content file(s) only after successfully serializing the file
          'Commented out - RKS 1/21/2009
          'DeleteAllContentFiles()
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Function TrimValue(ByVal lpValue As String,
                           ByVal lpTrimType As TrimPropertyType,
                           Optional ByVal lpTrimCharacter As String = " ") As String

      Dim lstrReturnValue As String

      ApplicationLogging.WriteLogEntry("Enter Document::TrimValue", TraceEventType.Verbose)
      Try
        Select Case lpTrimType
          Case TrimPropertyType.Left
            lstrReturnValue = lpValue.TrimStart(lpTrimCharacter)
          Case TrimPropertyType.Right
            lstrReturnValue = lpValue.TrimEnd(lpTrimCharacter)
          Case TrimPropertyType.Both
            lstrReturnValue = lpValue.Trim(lpTrimCharacter)
          Case Else
            lstrReturnValue = lpValue
        End Select

        ApplicationLogging.WriteLogEntry("Exit Document::TrimValue", TraceEventType.Verbose)
        Return lstrReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return ""
      End Try

    End Function

    Private Function GetDocumentPath() As String

      Dim lstrContentPathParts() As String
      Dim lstrDocumentPath As String = String.Empty

      ApplicationLogging.WriteLogEntry("Enter Document::GetDocumentPath", TraceEventType.Verbose)
      Try
        ' Get the ContentPath directories
        If Versions.Count > 0 AndAlso FirstVersion.HasContent = True AndAlso
  Not String.IsNullOrEmpty(FirstVersion.PrimaryContent.ContentPath) Then

          lstrContentPathParts = FirstVersion.PrimaryContent.ContentPath.Split("\")

          For lintCounter As Integer = 0 To lstrContentPathParts.Length - 1
            lstrDocumentPath &= lstrContentPathParts(lintCounter) & "\"
            If lstrContentPathParts(lintCounter) = Me.ID Then
              Exit For
            End If
          Next

          ApplicationLogging.WriteLogEntry("Exit Document::GetDocumentPath", TraceEventType.Verbose)

          Return lstrDocumentPath

        Else
          Return String.Empty
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return ""
      End Try

    End Function

    Public Sub AddFolderPath(ByVal lpFolderPath As String)
      Try
        FolderPathsProperty.Values.Add(New Value(lpFolderPath))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <JsonIgnore()>
    Public ReadOnly Property FolderPathsProperty() As ECMProperty
      Get
        Try

          If mobjFolderPathsProperty Is Nothing Then
            mobjFolderPathsProperty = GetFolderPathProperty()
          End If

          Return mobjFolderPathsProperty

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Private Function GetFolderPathsProperty() As ECMProperty
    '  Try

    '    Dim lobjFolderPathsProperty As ECMProperty

    '    lobjFolderPathsProperty = Properties(FOLDER_PATHS_PROPERTY_NAME)
    '    If lobjFolderPathsProperty IsNot Nothing Then
    '      Return lobjFolderPathsProperty
    '    End If

    '    lobjFolderPathsProperty = Properties(FOLDER_PATH_PROPERTY_NAME)
    '    If lobjFolderPathsProperty IsNot Nothing Then
    '      Return lobjFolderPathsProperty
    '    Else
    '      Dim lobjFoldersFiledIn As New ECMProperty(PropertyType.ecmString, "Folders Filed In", Cardinality.ecmMultiValued)
    '      Properties.Add(lobjFoldersFiledIn)
    '      Return lobjFoldersFiledIn
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Private Function GetFolderPaths(Optional ByVal lpAddBasePath As String = "",
                                Optional ByVal lpAddPathLocation As ePathLocation = ePathLocation.Front) As ObservableCollection(Of String) 'Collections.Specialized.StringCollection

      'Dim lobjFolderPaths As New Collections.Specialized.StringCollection
      Dim lobjFolderPath As Object
      Dim lstrFolderPath As String
      Dim lobjFolderPaths As New ObservableCollection(Of String)
      Dim lobjFolderPathsProperty As ECMProperty

      ApplicationLogging.WriteLogEntry("Enter Document::GetFolderPaths", TraceEventType.Verbose)
      Try
        lobjFolderPathsProperty = GetFolderPathProperty()
        If lobjFolderPathsProperty.Values Is Nothing Then
          lobjFolderPath = lobjFolderPathsProperty.Value
          If TypeOf lobjFolderPath Is Value Then
            lstrFolderPath = lobjFolderPath.Value
          Else
            lstrFolderPath = lobjFolderPath
          End If
          If lpAddPathLocation = ePathLocation.Front Then
            ' lobjFolderPaths.Add(lpAddBasePath & lobjFolderPathsProperty.Value.ToString)
            lobjFolderPaths.Add(String.Format("{0}{1}", lpAddBasePath, lstrFolderPath))
          Else
            ' lobjFolderPaths.Add(lobjFolderPathsProperty.Value.ToString & lpAddBasePath)
            lobjFolderPaths.Add(String.Format("{1}{0}", lpAddBasePath, lstrFolderPath))
          End If
          ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPaths", TraceEventType.Verbose)
          Return lobjFolderPaths
        Else
          Dim lintFolderCount As Integer = lobjFolderPathsProperty.Values.Count - 1
          For lintFolderCounter As Integer = 0 To lintFolderCount
            lobjFolderPath = lobjFolderPathsProperty.Values.Item(lintFolderCounter)
            If TypeOf lobjFolderPath Is Value Then
              lstrFolderPath = lobjFolderPath.Value
            Else
              lstrFolderPath = lobjFolderPath
            End If
            If lpAddPathLocation = ePathLocation.Front Then
              ' lobjFolderPaths.Add(lpAddBasePath & lobjFolderPathsProperty.Values.Item(lintFolderCounter).ToString)
              lobjFolderPaths.Add(String.Format("{0}{1}", lpAddBasePath, lstrFolderPath))
            Else
              ' lobjFolderPaths.Add(lobjFolderPathsProperty.Values.Item(lintFolderCounter).ToString & lpAddBasePath)
              lobjFolderPaths.Add(String.Format("{1}{0}", lpAddBasePath, lstrFolderPath))
            End If
          Next
          ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPaths", TraceEventType.Verbose)
          Return lobjFolderPaths
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPaths", TraceEventType.Verbose)
        Return New ObservableCollection(Of String) 'Collections.Specialized.StringCollection
      End Try

    End Function

    Public Function GetPrimaryFolderPath() As String
      Try
        Dim lobjPrimaryFolderPath As Object
        Dim lstrPrimaryFolderPath As String = String.Empty

        If FolderPathsProperty IsNot Nothing AndAlso FolderPathsProperty.HasValue Then

          If FolderPathsProperty.Cardinality = Cardinality.ecmSingleValued Then
            lstrPrimaryFolderPath = FolderPathsProperty.ValueString
          Else
            lobjPrimaryFolderPath = CType(FolderPathsProperty, MultiValueStringProperty).Values.FirstOrDefault()
            If TypeOf lobjPrimaryFolderPath Is Value Then
              lstrPrimaryFolderPath = DirectCast(lobjPrimaryFolderPath, Value).ToString()
            Else
              lstrPrimaryFolderPath = lobjPrimaryFolderPath.ToString()
            End If
            'lstrPrimaryFolderPath = CType(FolderPathsProperty, MultiValueStringProperty).Values.FirstOrDefault()
          End If

        End If

        Return lstrPrimaryFolderPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function CreateAppendedFolderPath(lpPathToAppend As String, lpFolderDelimiter As String) As String
      Try

        Dim lstrAppendedFolderPath As String = String.Empty
        Dim lstrPrimarySourceFolderPath As String = GetPrimaryFolderPath()

        If Not String.IsNullOrEmpty(lstrPrimarySourceFolderPath) Then
          If lstrPrimarySourceFolderPath.Length > 2 Then
            If ((lstrPrimarySourceFolderPath.Substring(1, 1) = ":") AndAlso
    (lstrPrimarySourceFolderPath.Substring(2, 1) = "\")) Then
              ' This is a file system folder path using a drive letter, 
              ' we don't want to insert the drive letter into this product path.
              lstrPrimarySourceFolderPath = lstrPrimarySourceFolderPath.Remove(0, 3)
            End If
          End If

          lstrPrimarySourceFolderPath = lstrPrimarySourceFolderPath.Replace("/", lpFolderDelimiter)

          'If lstrDestinationFolderPath.Contains(String.Format("{0}{1}", lpFolderDelimiter, PRIMARY_SOURCE_FOLDER_PATH)) AndAlso lstrPrimarySourceFolderPath.StartsWith(lpFolderDelimiter) Then
          '  lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(PRIMARY_SOURCE_FOLDER_PATH, lstrPrimarySourceFolderPath.TrimStart(lpFolderDelimiter))
          'ElseIf lstrPrimarySourceFolderPath.StartsWith(lpFolderDelimiter) Then
          '  lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(PRIMARY_SOURCE_FOLDER_PATH, lstrPrimarySourceFolderPath)
          'Else
          '  lstrDestinationFolderPath = lstrDestinationFolderPath.Replace(PRIMARY_SOURCE_FOLDER_PATH, String.Format("{0}{1}", lpFolderDelimiter, lstrPrimarySourceFolderPath))
          'End If

          lstrAppendedFolderPath = String.Format("{0}{1}{2}", lpPathToAppend.TrimEnd(lpFolderDelimiter), lpFolderDelimiter, lstrPrimarySourceFolderPath.TrimStart(lpFolderDelimiter))

        Else
          lstrAppendedFolderPath = lpPathToAppend
        End If

        Return lstrAppendedFolderPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetFolderPathProperty() As ECMProperty
      Try

        Dim lobjFolderPathProperty As ECMProperty = Nothing

        ' Try using the name 'Folders Filed In'
        lobjFolderPathProperty = If(If(Properties(FOLDER_PATHS_PROPERTY_NAME), Properties(FOLDER_PATH_PROPERTY_NAME)), Properties(FOLDERPATH_PROPERTY_NAME))

        If lobjFolderPathProperty Is Nothing Then
          ' We did not find the property yet
          ' Just create it as 'Folders Filed In'
          'lobjFolderPathProperty = New ECMProperty(PropertyType.ecmString, "Folders Filed In", Cardinality.ecmMultiValued)
          lobjFolderPathProperty = PropertyFactory.Create(PropertyType.ecmString, FOLDER_PATHS_PROPERTY_NAME, FOLDER_PATHS_PROPERTY_NAME, Cardinality.ecmMultiValued)
          Properties.Add(lobjFolderPathProperty)
        End If

        Return lobjFolderPathProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetFolderPathArray(ByVal lpPathFactory As PathFactory) As String()

      'Return GetFolderPathArray(lpPathFactory.BaseFolderPath, lpPathFactory.BasePathLocation, _
      '  lpPathFactory.Delimiter, lpPathFactory.LeadingDelimiter)

      Dim lobjFolderPaths As New ArrayList
      Dim lstrFolderPath As String
      Dim lobjFolderPathValues As Values

      ApplicationLogging.WriteLogEntry("Enter Document::GetFolderPathArray", TraceEventType.Verbose)
      Try

        Dim lobjFolderPathProperty As ECMProperty = GetFolderPathProperty()

        If lobjFolderPathProperty Is Nothing Then
          ' We could not find the folder path property
          Return Array.Empty(Of String)()
        End If

        Select Case lobjFolderPathProperty.Cardinality
          Case Cardinality.ecmSingleValued
            ' We have a single folder path
            lpPathFactory.SourceFolderPath = lobjFolderPathProperty.Value
            lstrFolderPath = lpPathFactory.CreateFolderPath
            ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
            Return New String(0) {lstrFolderPath}

          Case Else
            lobjFolderPathValues = lobjFolderPathProperty.Values
            If lobjFolderPathValues.Count = 0 Then
              If lobjFolderPathProperty.Value Is Nothing Then
                ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
                Return Array.Empty(Of String)()
              Else
                lpPathFactory.SourceFolderPath = lobjFolderPathProperty.Value
                lstrFolderPath = lpPathFactory.CreateFolderPath
                ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
                Return New String(0) {lstrFolderPath}
              End If
            Else
              lobjFolderPaths.Capacity = lobjFolderPathValues.Count - 1
            End If
            Dim lintFolderCount As Integer = lobjFolderPathValues.Count - 1
            For lintFolderCounter As Integer = 0 To lintFolderCount
              lpPathFactory.SourceFolderPath = lobjFolderPathValues.Item(lintFolderCounter)
              lobjFolderPaths.Add(lpPathFactory.CreateFolderPath)
            Next

            ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
            Return lobjFolderPaths.ToArray(GetType(String))
        End Select



        If Properties.PropertyExists(FOLDER_PATHS_PROPERTY_NAME) = False Then
          If Properties.PropertyExists(FOLDER_PATH_PROPERTY_NAME) = False Then
            Return Array.Empty(Of String)()
          Else
            ' We have a single folder path
            lpPathFactory.SourceFolderPath = Properties(FOLDER_PATH_PROPERTY_NAME).Value
            lstrFolderPath = lpPathFactory.CreateFolderPath
            ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
            Return New String(0) {lstrFolderPath}
          End If
        End If

        'lobjFolderPathValues = Properties(FOLDER_PATHS_PROPERTY_NAME).Values
        'If lobjFolderPathValues.Count = 0 Then
        '  If Properties(FOLDER_PATH_PROPERTY_NAME).Value Is Nothing Then
        '    ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
        '    Return New String() {}
        '  Else
        '    lpPathFactory.SourceFolderPath = Properties(FOLDER_PATH_PROPERTY_NAME).Value
        '    lstrFolderPath = lpPathFactory.CreateFolderPath
        '    ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
        '    Return New String(0) {lstrFolderPath}
        '  End If
        'Else
        '  lobjFolderPaths.Capacity = lobjFolderPathValues.Count - 1
        'End If
        'lintFolderCount = lobjFolderPathValues.Count - 1
        'For lintFolderCounter As Integer = 0 To lintFolderCount
        '  lpPathFactory.SourceFolderPath = lobjFolderPathValues.Item(lintFolderCounter)
        '  lobjFolderPaths.Add(lpPathFactory.CreateFolderPath)
        'Next

        'ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
        'Return lobjFolderPaths.ToArray(GetType(String))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::GetFolderPathArray", TraceEventType.Verbose)
        Return Array.Empty(Of String)()
      End Try

    End Function

    '' ''' <summary>
    '' ''' Recurses through each Version and Content and sets the parent Document reference
    '' ''' </summary>
    '' ''' <remarks></remarks>
    ''Private Sub SetDocumentReferences()
    ''  Try
    ''    If Versions IsNot Nothing Then
    ''      Versions.SetDocument(Me)
    ''      For Each lobjVersion As Version In Versions
    ''        lobjVersion.SetDocument(Me)
    ''        For Each lobjContent As Content In lobjVersion.Contents
    ''          lobjContent.SetVersion(lobjVersion)
    ''        Next
    ''      Next
    ''    End If
    ''  Catch ex As Exception
    ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''    '  Re-throw the exception to the caller
    ''    Throw
    ''  End Try
    ''End Sub

    Private Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
      ApplicationLogging.WriteLogEntry("Enter Document::LoadFromXmlDocument", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lobjdocument As Document = DeSerialize(lpXML)
        Helper.AssignObjectProperties(lobjdocument, Me)

        SetDocumentReferences()

        ' We are checking to see if the existing 'cdf' file has a header
        ' If it does not have a header we want to skip over the update header
        If lobjdocument.HeaderString.Length > 0 Then
          '  Recreate the header
          UpdateHeader()
        End If

        ' ''  Recreate the header from the header string
        ''Dim lstrHeaderXML As String = Core.Password.EscapedDecrypt(HeaderString)

        ''Dim lobjTempHeader As New Header
        ''mobjHeader = Serializer.Deserialize.XmlString(lstrHeaderXML, lobjTempHeader.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally
        ApplicationLogging.WriteLogEntry("Exit Document::LoadFromXmlDocument", TraceEventType.Verbose)
      End Try
    End Sub

    '''' <summary>
    '''' Updates the RelativePath of all the child Content objects
    '''' </summary>
    '''' <param name="lpCurrentDocumentPath">The path for this document's CDF file</param>
    '''' <remarks></remarks>
    'Private Sub SetRelativeContentPaths(ByVal lpCurrentDocumentPath As String)

    '  Dim lstrContentPath As String = ""
    '  'Dim lstrContentDir As String
    '  Dim lstrCDFDir As String

    '  ApplicationLogging.WriteLogEntry("Enter Document::SetRelativeContentPaths", TraceEventType.Verbose)
    '  Try

    '    lstrCDFDir = Path.GetDirectoryName(lpCurrentDocumentPath)

    '    ' Iterate through each content element of each version and set the RelativePath
    '    For Each lobjVersion As Version In Me.Versions
    '      For Each lobjContent As Content In lobjVersion.Contents
    '        ' Get the ContentPath
    '        lstrContentPath = lobjContent.ContentPath
    '        ' Set the RelativePath 
    '        lobjContent.RelativePath = RelativePath.GetRelativePath(lstrCDFDir, lstrContentPath, "\")
    '      Next
    '    Next

    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, String.Format("Document::SetRelativeContentPaths('{0}')", lpCurrentDocumentPath))
    '    '  Re-throw the exception to the caller
    '    Throw
    '  Finally
    '    ApplicationLogging.WriteLogEntry("Exit Document::SetRelativeContentPaths", TraceEventType.Verbose)
    '  End Try

    'End Sub

    ''' <summary>
    ''' Gets a reference to the ContentSource property
    ''' </summary>
    ''' <returns>The current document instance ContentSource object reference</returns>
    ''' <remarks>Used to check the availability of the ContentSource property for instance methods depending on a valid ContentSource reference</remarks>
    Protected Function GetContentSource() As ContentSource
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ' Make sure we have a valid ContentSource object reference

        ' It can't be a null reference
        If ContentSource Is Nothing Then
          Throw New InvalidContentSourceException
        End If

        ' The Provider reference can't be null either
        If ContentSource.Provider Is Nothing Then
          Throw New InvalidContentSourceException("The content source has no provider available to perform this operation")
        End If

        Return ContentSource

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Function GetBcsProvider() As IBasicContentServicesProvider
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If mobjRMProvider Is Nothing Then

          ' The Provider has to implement the BCS interface
          If GetContentSource.Provider.SupportsInterface(ProviderClass.BasicContentServices) = False Then
            Throw New InvalidContentSourceException(String.Format(
                                                                  "The content source '{0}' provider '{1}' does not implement the IBasicContentServicesProvider interface",
                                                                  ContentSource.Name, ContentSource.ProviderName))
          End If

          ' Set the BcsProvider reference
          mobjRMProvider = CType(Me.ContentSource.Provider, IBasicContentServicesProvider)

        End If

        ' Return the BcsProvider reference
        Return mobjRMProvider

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Checks to see if the current content source provider implements IRecordsManager.  
    ''' If it does then it returns a reference to the provider.
    ''' </summary>
    ''' <returns>An object reference to the current content source provider.</returns>
    ''' <remarks>If the ContentSource property of the document is not initialized or 
    ''' the provider used by the content source does not implement the IRecordsManager 
    ''' interface, an InvalidContentSourceException will be thrown.</remarks>
    Protected Function GetRMProvider() As Object
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If mobjRMProvider Is Nothing Then


          ' The Provider has to implement the RM interface
          If GetContentSource.Provider.SupportsInterface(ProviderClass.RecordsManager) = False Then
            Throw New InvalidContentSourceException(String.Format(
                                                                  "The content source '{0}' provider '{1}' does not implement the IRecordsManager interface",
                                                                  ContentSource.Name, ContentSource.ProviderName))
          End If

          ' Set the RMProvider reference
          ' Since the IRecordsManager interface is defined in the Ecmg.Cts.Records dll 
          ' and we do not have a reference to it here
          ' We will return an object reference
          mobjRMProvider = Me.ContentSource.Provider

        End If

        ' Return the RMProvider reference
        Return mobjRMProvider

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Deletes all the content files for the document
    ''' </summary>
    ''' <remarks>Use with caution.  
    ''' This method is only intended for use after successfully serializing the document with a non-reference StorageType.  
    ''' Otherwise there is no way to retrieve the content afterwards.</remarks>
    Private Sub DeleteAllContentFiles()
      Try
        For Each lobjVersion As Version In Me.Versions
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContent.DeleteContentFiles()
          Next
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Initializes the document hash
    ''' </summary>
    ''' <returns>Always return an empty string.</returns>
    ''' <remarks>Deprecated</remarks>
    Private Function GenerateHash() As String
      Try

        Return String.Empty

        ''TODO: somehow generate the hash without running the document xml into a string
        'If Me.StorageType = Content.StorageTypeEnum.EncodedCompressed Or Me.StorageType = Content.StorageTypeEnum.EncodedUnCompressed Then
        '  Return ""
        'End If

        '' This is to make sure the method is called only once
        ''Static lblnIsFirstCall As Boolean = True

        ''  Dim lstrHash As String

        ' '' Check to see if the Hash value has already been set.  If so then do not recalculate it.
        ''If mstrHash.Length > 0 OrElse lblnIsFirstCall = False Then
        ''  Return mstrHash
        ''End If

        '' Set the first call flag so that it will not get to this point a second time
        ''lblnIsFirstCall = False

        '' Create a stream to use for calculating the hash
        '' Dim lobjStream As Stream = New MemoryStream

        ''Dim lstrHex As String = Core.Password.Encrypt(Me.ToString).Hex
        'Dim lobjData As Byte()
        'Dim lstrDocumentXML As String

        'If mstrHeaderString.Length > 0 Then
        '  Dim lstrHeaderString As String = mstrHeaderString
        '  mstrHeaderString = ""
        '  lstrDocumentXML = Me.ToXmlString
        '  mstrHeaderString = lstrHeaderString
        'Else
        '  lstrDocumentXML = Me.ToXmlString
        'End If
        ''#If DEBUG Then
        ''          Dim lstrMe As String = Me.ToString
        ''          Debug.WriteLine(String.Format("Me with HeaderString: {0},{1}", lstrMe, vbCrLf))
        ''#End If
        ''lobjData = Encoding.UTF8.GetBytes(Me.ToString)
        ''mstrHeaderString = lstrHeaderString
        ''Else
        ' ''#If DEBUG Then
        ' ''          Dim lstrMe As String = Me.ToString
        ' ''          Debug.WriteLine(String.Format("Me with without HeaderString: {0},{1}", lstrMe, vbCrLf))
        ' ''#End If
        ''lobjData = Encoding.UTF8.GetBytes(Me.ToString)
        ''End If

        'lobjData = Encoding.UTF8.GetBytes(lstrDocumentXML)

        '' Initialize a byte array to fill the stream with


        '' Write the byte array to the stream
        ''lobjStream.Write(lobjData, 0, lobjData.Length)

        '' Create a Hash object to calculate the hash
        ''Dim lobjHash As New Utilities.Encryption.Hash(Utilities.Encryption.Hash.Provider.SHA256)

        ''lblnIsFirstCall = True

        'Return HashBytes(lobjData)

        ' '' Calculate the hash and get the text value
        ''Dim _Hash As New System.Security.Cryptography.SHA256Managed
        ''Dim b As Byte() = _Hash.ComputeHash(lobjData)
        ''lstrHash = System.Text.Encoding.UTF8.GetString(b)
        ' ''mstrHash = lobjHash.Calculate(lobjStream).Text

        ' '' Encrypt the hash value and store the HEX representation
        ''lstrHash = Core.Password.Encrypt(lstrHash).Hex

        ' '' Encrypt and URL escape the hash value
        ' ''mstrHash = Helper.EscapedEncrypt(mstrHash)

        ' '' Reset the flag
        ''lblnIsFirstCall = True

        ''Return lstrHash

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Protected Friend Shared Function HashBytes(ByVal lpBytesToHash() As Byte) As String
      Try
        Dim lstrHash As String

        ' Calculate the hash and get the text value
        Dim _Hash As New System.Security.Cryptography.SHA256Managed
        Dim b As Byte() = _Hash.ComputeHash(lpBytesToHash)
        lstrHash = System.Text.Encoding.UTF8.GetString(b)
        'mstrHash = lobjHash.Calculate(lobjStream).Text

        ' Encrypt the hash value and store the HEX representation
        lstrHash = Core.Password.Encrypt(lstrHash).Hex

        Return lstrHash
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the document has been modified since 
    ''' its original creation or the last transformation, whichever came last.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>True if the document has not been altered.  False if the document has been altered.</remarks>
    Public Function CheckHash() As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return CheckHash(HashLevel.Last)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the document has been modified since 
    ''' its original creation or the last transformation, depending on the specified HashLevel.
    ''' </summary>
    ''' <param name="lpHashLevel">Original or Last</param>
    ''' <returns></returns>
    ''' <remarks>True if the document has not been altered.  False if the document has been altered.</remarks>
    Public Function CheckHash(ByVal lpHashLevel As HashLevel) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lstrHashToCheck As String

        ' If we have no document header then we have nothing to check against, so we need to return false
        If Me.Header Is Nothing Then Return False

        Select Case lpHashLevel
          Case HashLevel.Last
            lstrHashToCheck = Me.Header.GetLatestHash()

          Case HashLevel.Original
            lstrHashToCheck = Me.Header.OriginalHash

          Case Else
            ' If they gave us an invalid hash level we will give them the latest by default
            lstrHashToCheck = Me.Header.GetLatestHash()
        End Select

        Dim lstrCurrentHash As String = GenerateHash()

        Return String.Equals(lstrCurrentHash, lstrHashToCheck)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates or updates the document header as necessary
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If you need access to the document header and it is null, then call this method to initialize it.</remarks>
    Friend Sub UpdateHeader()
      Try

        '  Create the header if it does not already exist
        If mobjHeader Is Nothing Then
          'If lpEncryptedHeaderString.Length = 0 Then
          If mstrHeaderString.Length = 0 Then
            If mstrHash.Length = 0 Then
              ' The hash has not yet been created.
              '   If this method was called from a Serialize method then we will 
              '   generate a new hash as it can be assumed that this is a newly created document.
              '   
              '   If this method was called from a Deserialize Method then we can assume
              '   that either the document was originally created before the header mechanism
              '   was implemented or that the original header has been damaged.
              '       In that case we will not generate a new hash but we will instead 
              '       set the value for the hash to 'NO_ORIGINAL_HASH'.
              '
              '     This will allow us to combat a situation where a malicious user might 
              '     damage or destroy the header in an attempt to have a new hash created.
              '     If we were to allow this, then the ability to prove a document has not been
              '     altered would be lost.
              '
              '     The original hash should never be re-created for any document
              If Helper.CallStackContainsMethodName("Deserialize") Then
                mstrHash = DocumentHeader.NO_ORIGINAL_HASH
              End If
              If Helper.CallStackContainsMethodName("Serialize", "Transform", "ToArchiveStream") Then
                mstrHash = GenerateHash()
              End If
            End If
            ' Create the header
            mobjHeader = New DocumentHeader(Me.ID, Me.SerializationPath, mstrHash, Me.ContentSource)
            ' Create the header string
            'mstrHeaderString = Core.Password.Encrypt(mobjHeader.ToString).Hex
            'mstrHeaderString = Core.Password.Encrypt(Helper.Compression.CompressString(mobjHeader.ToString)).Hex
            mstrHeaderString = mobjHeader.GenerateHeaderString
          Else
            'Dim lstrHeaderXML As String
            Try
              '  Reconstruct the header
              'lstrHeaderXML = Core.Password.EscapedDecryptFromHex(mstrHeaderString)
              ' Create the header
              'mobjHeader = New Header(lstrHeaderXML)
              mobjHeader = DocumentHeader.GenerateHeader(mstrHeaderString)
              mstrHash = Header.OriginalHash
              SerializationPath = Header.SerializationPath
            Catch FormatEx As FormatException
              ApplicationLogging.LogException(FormatEx, Reflection.MethodBase.GetCurrentMethod)
              ' We were unable to reconstruct the header from the header string
              ' We need to re-create it
              mobjHeader = New DocumentHeader(Me.ID, Me.SerializationPath, DocumentHeader.NO_ORIGINAL_HASH, Me.ContentSource)
              ' Create the header string
              'mstrHeaderString = Core.Password.Encrypt(mobjHeader.ToString).Hex
              mstrHeaderString = mobjHeader.GenerateHeaderString
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              '  Re-throw the exception to the caller
              Throw
            End Try
          End If
        Else
          ' Make sure the header string always reflects the current state of the header
          'mstrHeaderString = Core.Password.Encrypt(mobjHeader.ToString).Hex
          mstrHeaderString = mobjHeader.GenerateHeaderString
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    'Protected Friend Overridable Function DebuggerIdentifier() As String
    '  Dim lobjIdentifierBuilder As New StringBuilder
    '  Try
    '    Dim lstrName As String = Name

    '    lobjIdentifierBuilder.AppendFormat("{0} - ", ID)

    '    ' 1st show the class, if assigned.
    '    If DocumentClass.Length > 0 Then
    '      lobjIdentifierBuilder.AppendFormat("({0}) ", DocumentClass)
    '    End If

    '    If lstrName Is Nothing OrElse lstrName.Length = 0 Then
    '      lobjIdentifierBuilder.Append("Document identifier not set")
    '    Else
    '      lobjIdentifierBuilder.AppendFormat("{0}", lstrName)
    '    End If

    '    If Versions Is Nothing OrElse Versions.Count = 0 Then
    '      lobjIdentifierBuilder.Append(": No Versions")
    '    ElseIf Versions.Count = 1 Then
    '      lobjIdentifierBuilder.AppendFormat(": 1 version ({0})", Me.GetFirstVersion.DebuggerIdentifier)
    '    ElseIf Versions.Count > 1 Then
    '      lobjIdentifierBuilder.AppendFormat(": {0} versions 1st:({1})", Versions.Count, _
    '                                         Me.GetFirstVersion.DebuggerIdentifier)
    '    End If

    '    Dim lintInvalidRelationshipCount As Integer = InvalidRelationshipCount()

    '    If lintInvalidRelationshipCount > 0 Then
    '      lobjIdentifierBuilder.AppendFormat(" * {0} Relationship(s), {1} Invalid", Relationships.Count, InvalidRelationshipCount())
    '      'ElseIf lintInvalidRelationshipCount = 1 Then
    '      '  lobjIdentifierBuilder.AppendFormat(" * {0} Relationships, 1 Invalid", Relationships.Count)
    '    ElseIf Relationships IsNot Nothing AndAlso Relationships.Count > 0 Then
    '      lobjIdentifierBuilder.AppendFormat(" * {0} Relationship(s)", Relationships.Count)
    '    End If

    '    Return lobjIdentifierBuilder.ToString

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Return lobjIdentifierBuilder.ToString
    '  End Try
    'End Function

#Region "Verify Properties"

    Private Sub VerifyProperties(ByVal lpXmlDocument As XmlDocument)
      Try

        Dim lobjPropertyNodeList As XmlNodeList
        Dim lstrVersionXPath As String = String.Empty
        Dim lstrContentXPath As String = String.Empty
        Dim lobjContentsNodeList As XmlNodeList = Nothing
        Dim lobjContentNode As XmlNode = Nothing
        Dim lintContentCounter As Integer = 0

        ' Look through each document property, then each version property and content property

        ' Look through each document property
        lobjPropertyNodeList = lpXmlDocument.SelectNodes("Document/Properties/ECMProperty")
        VerifyProperties(lobjPropertyNodeList, Me)

        ' Look through each Version
        For Each lobjVersion As Version In Versions
          lstrVersionXPath = String.Format("Document/Versions/Version[@ID='{0}']", lobjVersion.ID)
          ' First look through the contents
          lstrContentXPath = String.Format("{0}/Contents", lstrVersionXPath)
          lobjContentsNodeList = lpXmlDocument.SelectNodes(lstrContentXPath)
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContentNode = lobjContentsNodeList(lintContentCounter).FirstChild
            lobjPropertyNodeList = lobjContentNode.SelectNodes("Metadata/ECMProperty")
            VerifyProperties(lobjPropertyNodeList, lobjContent)
          Next
          ' Now look through the version properties themselves
          lobjPropertyNodeList = lpXmlDocument.SelectNodes(String.Format("{0}/Properties/ECMProperty", lstrVersionXPath))
          VerifyProperties(lobjPropertyNodeList, lobjVersion)
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Iterates through all properties and makes sure 
    ''' that the values are set from the specified source file.
    ''' </summary>
    ''' <remarks>This is necessary when deserializing from older 
    ''' style documents that were serialized before we created 
    ''' the strongly typed property subclasses.  Only runs for 
    ''' Documents created before 1.6.7.0</remarks>
    Private Sub VerifyProperties(ByVal lpSourceFilePath As String)
      Try

        If String.Compare(CTS_Version, "1.6.7.0") > 0 Then
          Exit Sub
        End If

        Dim lobjXmlDocument As New XmlDocument

        If String.Compare(Path.GetExtension(lpSourceFilePath), ".cdf") = 0 Then
          ' This is a cdf file
          lobjXmlDocument.Load(lpSourceFilePath)

          VerifyProperties(lobjXmlDocument)

        ElseIf String.Compare(Path.GetExtension(lpSourceFilePath), ".cpf") = 0 Then
          ' This is a cpf file
          Dim lobjZipFile As ZipFile = ZipFile.Read(lpSourceFilePath)
          Dim lobjDocumentStream As MemoryStream

          For Each lobjZipEntry As ZipEntry In lobjZipFile
            If lobjZipEntry.FileName.EndsWith("cdf") Then
              lobjDocumentStream = New MemoryStream
              lobjZipEntry.Extract(lobjDocumentStream)
              lobjDocumentStream.Position = 0
              lobjXmlDocument.Load(lobjDocumentStream)
              VerifyProperties(lobjXmlDocument)
            End If
          Next

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub VerifyProperties(ByVal lpXmlNodeList As XmlNodeList,
                                     ByVal lpMetaHolder As IMetaHolder)
      Try

        'Dim lstrIdAttribute As XmlAttribute = Nothing
        Dim lstrPropertyName As String = String.Empty
        Dim lobjPropertyValueNode As XmlNode = Nothing
        Dim lstrPropertyValue As String = String.Empty
        Dim lobjTargetProperty As ECMProperty = Nothing
        Dim lobjReplacementProperty As ECMProperty = Nothing

        For Each lobjPropertyNode As XmlNode In lpXmlNodeList
          'lstrIdAttribute = lobjPropertyNode.Attributes("ID")
          'lstrPropertyName = lstrIdAttribute.Value
          If lobjPropertyNode.HasChildNodes Then
            lstrPropertyName = lobjPropertyNode.Item("Name").InnerText
          End If
          lobjTargetProperty = lpMetaHolder.Metadata.Item(lstrPropertyName)
          If lobjTargetProperty Is Nothing Then
            Continue For
          End If
          If lobjTargetProperty.Cardinality = Cardinality.ecmSingleValued Then
            lobjPropertyValueNode = lobjPropertyNode.SelectSingleNode("Value")
            lstrPropertyValue = lobjPropertyValueNode.InnerText
            If Not String.IsNullOrEmpty(lobjPropertyValueNode.InnerText) Then
              lobjTargetProperty.Value = lobjPropertyValueNode.InnerText
            End If
          Else
            lobjPropertyValueNode = lobjPropertyNode.SelectSingleNode("Values")
            If lobjPropertyValueNode IsNot Nothing Then
              For Each lobjMVValueNode As XmlNode In lobjPropertyValueNode.ChildNodes
                If Not String.IsNullOrEmpty(lobjMVValueNode.InnerText) Then
                  If lobjTargetProperty.Values.Contains(lobjMVValueNode.InnerText) = False Then
                    lobjTargetProperty.Values.Add(lobjMVValueNode.InnerText)
                  End If
                End If
              Next
            End If
          End If
          lobjReplacementProperty = PropertyFactory.Create(lobjTargetProperty)
          lpMetaHolder.Metadata.Replace(lstrPropertyName, lobjReplacementProperty)
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Private Sub MobjFolderPaths_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles MobjFolderPaths.CollectionChanged
      Try
        Select Case e.Action
          Case Specialized.NotifyCollectionChangedAction.Reset
            FolderPathsProperty.Clear()
          Case Specialized.NotifyCollectionChangedAction.Add
            ' FolderPathsProperty.Values.AddRange(e.NewItems)
            For Each lstrFolderPath As String In e.NewItems
              FolderPathsProperty.AddValue(lstrFolderPath, False)
            Next
          Case Specialized.NotifyCollectionChangedAction.Remove
            FolderPathsProperty.Values.Remove(e.OldItems)
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub CleanTempFiles()
      Try
        For Each lobjTempFileStream As FileStream In mobjTempFileStreams
          lobjTempFileStream.Dispose()
        Next

        For Each lstrTempFilePath As String In mobjTempFilePaths
          If IO.File.Exists(lstrTempFilePath) Then
            IO.File.Delete(lstrTempFilePath)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable Implementation"

    ''' <summary>
    ''' Generates a deep copy of the document
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Clone() As Object Implements System.ICloneable.Clone
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ' Instead of trying to individually reproduce each element of the document it is 
        ' more reliable to serialize the document to a stream and reconstruct it.

        'With lobjDocument
        '  If Me.ContentSource IsNot Nothing Then
        '    .ContentSource = New ContentSource(Me.ContentSource.ConnectionString)
        '  End If
        '  .CTS_Version = Me.CTS_Version.Clone
        '  .DocumentClass = Me.DocumentClass.Clone
        '  .Hash = Me.Hash.Clone
        '  .HeaderString = Me.HeaderString.Clone
        '  .ID = Me.ID.Clone
        '  .Identifier = Me.Identifier.Clone
        '  .Metadata = Me.Metadata.Clone
        '  .ObjectID = Me.ObjectID.Clone
        '  .Properties = Me.Properties.Clone
        '  .Relationships = Me.Relationships
        '  .SerializationPath = Me.SerializationPath.Clone
        '  .StorageType = Me.StorageType
        '  .Versions = Me.Versions.Clone
        'End With

        Dim lobjDocumentStream As Stream = Me.ToArchiveStream
        'Return New Document(Me.Serialize)

        ' <Modified by: Ernie at 4/13/2012-10:14:12 AM on machine: ERNIE-M4400>
        ' Dim lobjDocument As New Document()
        Dim lobjDocument As New Document(lobjDocumentStream)

        ' </Modified by: Ernie at 4/13/2012-10:14:12 AM on machine: ERNIE-M4400>
        Return lobjDocument

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Clone(ByVal lpWithoutContent As Boolean) As Object
      Try
        If lpWithoutContent = False Then
          Return Me.Clone()
        Else
          Dim lobjDocument As Document = Nothing

          lobjDocument = New Document(Me.ToStream())
          lobjDocument.ClearAllContents()
          Return lobjDocument

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

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
            mobjDataTable?.Dispose()

            mobjFolderPathsProperty?.Dispose()

            ' Dispose of all document level properties
            mobjProperties?.Dispose()

            mobjRelationships?.Dispose()

            mobjRelatedDocuments?.Dispose()

            ' Dispose of all versions
            If Me.Versions IsNot Nothing Then
              For Each lobjVersion As Version In Me.Versions
                lobjVersion.Dispose()
              Next
            End If
          End If

          mobjProperties = Nothing
          mobjRelationships = Nothing
          mobjRelatedDocuments = Nothing

          mobjVersions = Nothing

          mobjFolderPathsProperty = Nothing

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' DISPOSETODO: set large fields to null.
          mobjContentSource = Nothing

          CleanTempFiles()

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

#Region "ILoggable Implementation"

    'Private mobjLogSession As Gurock.SmartInspect.Session

    'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
    '  Try
    '    ApplicationLogging.FinalizeLogSession(mobjLogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.LightYellow)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
    '  Get
    '    Try
    '      If mobjLogSession Is Nothing Then
    '        InitializeLogSession()
    '      End If
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

#End Region

  End Class

End Namespace
