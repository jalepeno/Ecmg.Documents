'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Exceptions
Imports Documents.Serialization
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace SerializationUtilities

#Region "Public Interfaces"

  ''' <summary>
  ''' Used by all objects which can be written to, or re-instantiated from xml files.
  ''' </summary>
  ''' <remarks>Primarily used in to serialize with object.Serialize and invoked inside object constructors which take an xml file path as a single constructor parameter.</remarks>
  Public Interface ISerialize
    Sub Serialize(ByVal lpFilePath As String)
    Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String)
    Sub Serialize(ByVal lpFilePath As String,
      ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String)
    Function Serialize() As Xml.XmlDocument
    Function ToXmlString() As String
    Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object
    Function DeSerialize(ByVal lpXML As Xml.XmlDocument) As Object
    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property DefaultFileExtension() As String
  End Interface

  ''' <summary>
  ''' Used by all objects that can be written to, or re-instantiated from streams.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IStreamSerialize
    Function SerializeToStream() As IO.Stream
    Function DeSerialize(ByVal lpStream As IO.Stream) As Object
  End Interface

  Public Interface IZipStreamSerialize
    Function Deserialize(ByVal lpStream As IO.Stream, ByVal lpZipFile As ZipFile) As Object
  End Interface

#End Region

  Public Enum SourceType
    File = 0
    Stream = 1
  End Enum

  Partial Public Class Serializer

    Partial Public Class Serialize

      ''' <summary>
      ''' Replace invalid characters (\ / : * ? " &lt; &gt; |) with _
      ''' </summary>
      ''' <param name="lpFileName">The file name to clean</param>
      ''' <param name="lpCleanChar">The character to replace with</param>
      ''' <returns>Cleaned up file name</returns>
      ''' <remarks></remarks>
      Shared Function CleanFile(ByVal lpFileName As String, ByVal lpCleanChar As Char) As String

        Try

          ' Eliminate all instances of \ / : * ? " < > |
          Dim lstrCleanFileName As String = lpFileName
          lstrCleanFileName = lstrCleanFileName.Replace("\", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("/", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace(":", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("*", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("?", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("""", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("<", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace(">", lpCleanChar)
          lstrCleanFileName = lstrCleanFileName.Replace("|", lpCleanChar)

          Return lstrCleanFileName

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try


      End Function

      ''' <summary>
      ''' Replace invalid characters (\ / : * ? " &lt; &gt; |) with _
      ''' </summary>
      ''' <param name="lpFileName">The file name to clean</param>
      ''' <returns>Cleaned up file name</returns>
      ''' <remarks>For more conrol over the replacement character, use CleanFile(lpFileName, lpCleanChar)</remarks>
      Shared Function CleanFile(ByVal lpFileName As String) As String
        Return CleanFile(lpFileName, "_")
      End Function

      ''' <summary>
      ''' 
      ''' </summary>
      ''' <param name="lpObject"></param>
      ''' <param name="lpFilePath"></param>
      ''' <param name="lpSchemaLocation"></param>
      ''' <param name="lpDeclaration"></param>
      ''' <param name="lpWriteProcessingInstruction"></param>
      ''' <param name="lpXSLPath"></param>
      ''' <param name="lpCleanChar"></param>
      ''' <param name="lpClean"></param>
      ''' <remarks></remarks>
      Shared Sub XmlFile(ByVal lpObject As Object,
        ByVal lpFilePath As String,
        Optional ByVal lpSchemaLocation As String = "",
        Optional ByVal lpDeclaration() As String = Nothing,
        Optional ByVal lpWriteProcessingInstruction As Boolean = False,
        Optional ByVal lpXSLPath As String = "",
        Optional ByVal lpCleanChar As Char = "_",
        Optional ByVal lpClean As Boolean = False,
        Optional ByVal lpCulture As CultureInfo = Nothing)

        Dim lstrInstructionName As String
        Dim lstrInstructionText As String

        Dim lobjWriter As EnhancedXmlTextWriter = Nothing

        Try

          ' ApplicationLogging.WriteLogEntry("Enter Serialize::XmlFile", TraceEventType.Verbose)

          '' Clean the file name
          'lpFilePath = CleanFile(lpFilePath, lpCleanChar)

          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpObject.GetType)
          lobjWriter = New EnhancedXmlTextWriter(lpFilePath, System.Text.Encoding.UTF8)

          If lpCulture IsNot Nothing Then
            lobjWriter.Culture = lpCulture
          End If

          ' Write a human readable file
          lobjWriter.Formatting = Formatting.Indented

          If Not lpDeclaration Is Nothing Then
            'Write the XML delcaration. 
            lobjWriter.WriteStartDocument()

            For Each lstrInstruction As String In lpDeclaration
              lobjWriter.WriteWhitespace(vbCrLf)
              lstrInstructionName = lstrInstruction.Split("^").GetValue(0)
              lstrInstructionText = lstrInstruction.Split("^").GetValue(1)
              lobjWriter.WriteProcessingInstruction(lstrInstructionName, lstrInstructionText)
            Next
            'Dim lstrProcessingInstruction As String = lpDeclaration '"type='text/xsl' href='book.xsl'"
            'lobjWriter.WriteProcessingInstruction("xml-stylesheet", lstrProcessingInstruction)
          End If

          If lpWriteProcessingInstruction = True Then

            'Write the ProcessingInstruction node.
            If lpWriteProcessingInstruction = True Then
              Dim lstrProcessingInstruction As String = "type=""text/xsl"" href=""" & lpXSLPath & """"
              lobjWriter.WriteProcessingInstruction("xml-stylesheet", lstrProcessingInstruction)
            End If

          End If

          lobjFormatter.Serialize(lobjWriter, lpObject)

        Catch IoEx As IOException
          If Helper.CallStackContainsMethodName("FindProvider") Then
            ' Ignore this error, it is most likely caused by multiple 
            ' threads loading up a provider at the same time.
          Else
            ApplicationLogging.LogException(IoEx, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New InvalidOperationException("Could not serialize object to xml. [" & ex.Message & "]", ex)
        Finally
          If Not lobjWriter Is Nothing Then lobjWriter.Close()

          If lpClean Then
            Helper.FormatXmlFile(lpFilePath)
          End If

          ' ApplicationLogging.WriteLogEntry("Exit Serialize::XmlFile", TraceEventType.Verbose)

          'If lpDeclaration.Length > 0 Then
          '  AddDeclaration(lpFilePath, lpDeclaration)
          'End If

        End Try

      End Sub

      Shared Sub CleanXmlFile(lpObject As Object, lpFilePath As String)
        Try
          XmlFile(lpObject, lpFilePath, String.Empty, Nothing, False, String.Empty, "_", True)
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Sub

      Shared Function Xml(ByVal lpObject As Object,
        Optional ByVal lpSchemaLocation As String = "",
        Optional ByVal lpDeclaration() As String = Nothing) As Xml.XmlDocument

        Dim lstrInstructionName As String
        Dim lstrInstructionText As String
        Dim lobjXMLStream As Stream

        Dim lobjWriter As XmlTextWriter = Nothing

        ' ApplicationLogging.WriteLogEntry("Enter Serialize::Xml", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpObject.GetType)
          'lobjWriter = New XmlTextWriter(lpFilePath, System.Text.Encoding.UTF8)
          lobjXMLStream = New System.IO.MemoryStream()


          If Not lpDeclaration Is Nothing Then
            ' We will use the XmlTextWriter
            lobjWriter = New XmlTextWriter(lobjXMLStream, System.Text.Encoding.UTF8)

            'Write the XML delcaration. 
            lobjWriter.WriteStartDocument()

            For Each lstrInstruction As String In lpDeclaration
              lobjWriter.WriteWhitespace(vbCrLf)
              lstrInstructionName = lstrInstruction.Split("^").GetValue(0)
              lstrInstructionText = lstrInstruction.Split("^").GetValue(1)
              lobjWriter.WriteProcessingInstruction(lstrInstructionName, lstrInstructionText)
            Next
            ''Write the ProcessingInstruction node.
            'Dim lstrProcessingInstruction As String = lpDeclaration '"type='text/xsl' href='book.xsl'"
            'lobjWriter.WriteProcessingInstruction("xml-stylesheet", lstrProcessingInstruction)

            ' Write a human readable file
            lobjWriter.Formatting = Formatting.Indented

            lobjFormatter.Serialize(lobjWriter, lpObject)

          Else
            ' Simply serialize straight to the stream          
            lobjFormatter.Serialize(lobjXMLStream, lpObject)
          End If

          Dim lobjXMLDocument As New XmlDocument
          If (lobjXMLStream.CanSeek) Then
            If (lobjXMLStream.Position <> 0) Then
              lobjXMLStream.Position = 0
            End If
          End If
          lobjXMLDocument.Load(lobjXMLStream)

          Return lobjXMLDocument

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New InvalidOperationException("Could not serialize object to xml. [" & ex.Message & "]", ex)
        Finally
          If Not lobjWriter Is Nothing Then lobjWriter.Close()
          ' ApplicationLogging.WriteLogEntry("Exit Serialize::Xml", TraceEventType.Verbose)

          'If lpDeclaration.Length > 0 Then
          '  AddDeclaration(lpFilePath, lpDeclaration)
          'End If

        End Try

      End Function

      Shared Sub AddComment(ByVal lpFilePath As String, ByVal lpComment As String)

        ' ApplicationLogging.WriteLogEntry("Exit Serialize::AddComment", TraceEventType.Verbose)
        Try

          Dim doc As New XmlDocument
          doc.Load(lpFilePath)
          'doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" & _
          '            "<title>Pride And Prejudice</title>" & _
          '            "</book>")

          'Create a comment.
          Dim newComment As XmlComment
          newComment = doc.CreateComment(lpComment)

          'Add the new node to the document.
          Dim root As XmlElement = doc.DocumentElement
          doc.InsertBefore(newComment, root)

          Console.WriteLine("Display the modified XML...")
          doc.Save(Console.Out)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit Serialize::AddComment", TraceEventType.Verbose)
        End Try

      End Sub

      Shared Sub AddDeclaration(ByVal lpFilePath As String, ByVal lpDeclaration As String)

        ' ApplicationLogging.WriteLogEntry("Enter Serialize::AddDeclaration", TraceEventType.Verbose)
        Try

          Dim doc As New XmlDocument
          doc.Load(lpFilePath)

          'Create an XML declaration. 
          Dim xmldecl As XmlDeclaration
          xmldecl = doc.CreateXmlDeclaration(lpDeclaration, Nothing, Nothing)

          'Add the new node to the document.
          Dim root As XmlElement = doc.DocumentElement
          doc.InsertBefore(xmldecl, root)

          Console.WriteLine("Display the modified XML...")
          doc.Save(Console.Out)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit Serialize::AddDeclaration", TraceEventType.Verbose)
        End Try

      End Sub

      Shared Function XmlElementString(ByVal lpObject As Object) As String
        Try
          Dim lstrXmlString As String = XmlString(lpObject)
          Dim lobjXmlDocument As New XmlDocument
          Dim lobjDocumentElement As XmlElement

          Dim lstrXSIRef As String = String.Format(" xmlns:xsi={0}http://www.w3.org/2001/XMLSchema-instance{0}", ControlChars.Quote)
          Dim lstrXSDRef As String = String.Format(" xmlns:xsd={0}http://www.w3.org/2001/XMLSchema{0}", ControlChars.Quote)

          '' Remove any instances of this namespace
          '' For some reason the CleanNameSpaceAttributes
          '' sometaimes fails to get this one
          'lstrXmlString = lstrXmlString.Replace(String.Format(" xmlns:xsi={0}http://www.w3.org/2001/XMLSchema-instance{0}", ControlChars.Quote), "")

          Try
            lobjXmlDocument.LoadXml(lstrXmlString)
          Catch XmlEx As XmlException
            If XmlEx.Message.StartsWith("'xsi' is an undeclared namespace.") Then
              lstrXmlString = InsertNameSpace(lstrXmlString, lstrXSIRef)
              lobjXmlDocument.LoadXml(lstrXmlString)
            End If
          End Try
          'lobjXmlDocument.LoadXml(lstrXmlString)
          lobjDocumentElement = lobjXmlDocument.DocumentElement

          ' Remove all the namespace attributes from all the elements
          'CleanNameSpaceAttributes(lobjDocumentElement)

          'Dim lstrReturnString As String = lobjXmlDocument.DocumentElement.OuterXml
          Dim lstrReturnString As String = Helper.FormatXmlString(lobjXmlDocument.DocumentElement.OuterXml)

          lstrReturnString = lstrReturnString.Replace(lstrXSIRef, String.Empty)
          lstrReturnString = lstrReturnString.Replace(lstrXSDRef, String.Empty)

          Return lstrReturnString

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Function

      Public Shared Sub CleanNameSpaceAttributes(ByRef lpElement As XmlElement)
        Try
          For lintAttributeCounter As Integer = lpElement.Attributes.Count - 1 To 0 Step -1
            If lpElement.Attributes(lintAttributeCounter).Name.StartsWith("xmlns:") Then
              lpElement.Attributes.Remove(lpElement.Attributes(lintAttributeCounter))
            End If
          Next

          For Each lobjNode As XmlNode In lpElement.ChildNodes
            If lobjNode.GetType.Name = "XmlElement" Then
              CleanNameSpaceAttributes(lobjNode)
            End If
          Next

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Sub

      'Shared Sub CustomFile(ByVal lpObject As Object, ByVal lpFilePath As String, ByVal lpFormatter As IFormatter)

      '  Dim lobjWriter As FileStream = Nothing
      '  ' ApplicationLogging.WriteLogEntry("Enter Serialize::CustomFile", TraceEventType.Verbose)
      '  Try

      '    '' Clean the file name
      '    'lpFilePath = CleanFile(lpFilePath)

      '    lobjWriter = File.Create(lpFilePath)
      '    lpFormatter.Serialize(lobjWriter, lpObject)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Finally
      '    If Not lobjWriter Is Nothing Then lobjWriter.Close()
      '    ' ApplicationLogging.WriteLogEntry("Exit Serialize::CustomFile", TraceEventType.Verbose)
      '  End Try

      'End Sub

      'Shared Function CustomString(ByVal lpObject As Object, ByVal lpFormatter As IFormatter) As String
      '  ' ApplicationLogging.WriteLogEntry("Enter Serialize::CustomString", TraceEventType.Verbose)
      '  Try
      '    Dim lobjBuffer As Byte() = CustomArray(lpObject, lpFormatter)
      '    Return Encoding.Default.GetString(lobjBuffer)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return ""
      '  Finally
      '    ApplicationLogging.WriteLogEntry("Exit Serialize::CustomString", TraceEventType.Verbose)
      '  End Try
      'End Function

      'Shared Function CustomArray(ByVal lpObject As Object, ByVal lpFormatter As IFormatter) As Byte()
      '  ApplicationLogging.WriteLogEntry("Enter Serialize::CustomArray", TraceEventType.Verbose)
      '  Try
      '    Return ToStream(lpObject, lpFormatter).ToArray
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return Nothing
      '  Finally
      '    ' ApplicationLogging.WriteLogEntry("Exit Serialize::CustomArray", TraceEventType.Verbose)
      '  End Try
      'End Function

      Shared Function XmlString(ByVal lpObject As Object) As String
        Dim lobjWriter As System.IO.StringWriter = Nothing

        'ApplicationLogging.WriteLogEntry("Enter Serialize::XmlString", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpObject.GetType)
          lobjWriter = New System.IO.StringWriter
          lobjFormatter.Serialize(lobjWriter, lpObject)
          Return lobjWriter.ToString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New Exception("Unable to Serialize to XMLString", ex)
        Finally
          If Not lobjWriter Is Nothing Then lobjWriter.Close()
          'ApplicationLogging.WriteLogEntry("Exit Serialize::XmlString", TraceEventType.Verbose)
        End Try
      End Function

      'Shared Function ToStream(ByVal lpObject As Object, ByVal lpFormatter As IFormatter) As MemoryStream
      '  Dim lobjWriter As MemoryStream = Nothing
      '  ' ApplicationLogging.WriteLogEntry("Enter Serialize::ToStream", TraceEventType.Verbose)
      '  Try
      '    lobjWriter = New MemoryStream
      '    lpFormatter.Serialize(lobjWriter, lpObject)
      '    If (lobjWriter.CanSeek) Then
      '      lobjWriter.Position = 0
      '    End If
      '    Return lobjWriter
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return Nothing
      '  Finally
      '    ' ApplicationLogging.WriteLogEntry("Exit Serialize::ToStream", TraceEventType.Verbose)
      '  End Try
      'End Function

      Shared Function ToStream(ByVal lpObject As Object) As MemoryStream
        Dim lobjWriter As MemoryStream = Nothing
        ApplicationLogging.WriteLogEntry("Enter Serialize::ToStream", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpObject.GetType)
          lobjWriter = New MemoryStream
          lobjFormatter.Serialize(lobjWriter, lpObject)
          If (lobjWriter.CanSeek) Then
            lobjWriter.Position = 0
          End If
          Return lobjWriter
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New SerializationException(String.Format("Unable to serialize {0} object to stream: {1}", lpObject.GetType.Name, ex.Message), ex)
        Finally
          ApplicationLogging.WriteLogEntry("Exit Serialize::ToStream", TraceEventType.Verbose)
        End Try
      End Function

      'Shared Function ToStream(ByVal lpObject As Object, ByVal lpFileName As String) As MemoryStream
      '  Dim lobjWriter As MemoryStream = Nothing
      '  ApplicationLogging.WriteLogEntry("Enter Serialize::ToStream", TraceEventType.Verbose)
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpObject.GetType)
      '    lobjWriter = New MemoryStream
      '    lobjFormatter.Serialize(lobjWriter, lpObject)
      '    If (lobjWriter.CanSeek) Then
      '      lobjWriter.Position = 0
      '    End If
      '    Return lobjWriter
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return Nothing
      '  Finally
      '    ApplicationLogging.WriteLogEntry("Exit Serialize::ToStream", TraceEventType.Verbose)
      '  End Try
      'End Function

      'Shared Function ToBinaryStream(ByVal lpObject As Object) As MemoryStream
      '  Dim lobjWriter As MemoryStream = Nothing
      '  ' ApplicationLogging.WriteLogEntry("Enter Serialize::ToStream", TraceEventType.Verbose)
      '  Try

      '    Dim lobjStream As New MemoryStream
      '    Dim lobjFormatter As IFormatter = New BinaryFormatter()

      '    lobjFormatter.Serialize(lobjStream, lpObject)
      '    If (lobjStream.CanSeek) Then
      '      lobjStream.Position = 0
      '    End If

      '    Return lobjStream

      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Return Nothing
      '  Finally
      '    ' ApplicationLogging.WriteLogEntry("Exit Serialize::ToStream", TraceEventType.Verbose)
      '  End Try
      'End Function

      'Shared Function SoapArray(ByVal lpObject As Object) As Byte()
      '  Return CustomArray(lpObject, New SoapFormatter)
      'End Function

      'Shared Sub SoapFile(ByVal lpObject As Object, ByVal lpFilePath As String)
      '  CustomFile(lpObject, lpFilePath, New SoapFormatter)
      'End Sub

      'Shared Function SoapString(ByVal lpObject As Object) As String
      '  Return CustomString(lpObject, New SoapFormatter)
      'End Function

      'Shared Function BinaryArray(ByVal lpObject As Object) As Byte()
      '  Return CustomArray(lpObject, New BinaryFormatter)
      'End Function

      'Shared Sub BinaryFile(ByVal lpObject As Object, ByVal lpFilePath As String)
      '  CustomFile(lpObject, lpFilePath, New BinaryFormatter)
      'End Sub

      Private Shared Function InsertNameSpace(lpXmlString As String, lstrXSIRef As String) As String
        Try
          ' Add the namespace to the root node somehow
          ' Perhaps with a regular expression?
          ''Dim lobjNameTable As New NameTable
          ''Dim lobjNamespaceManager As New XmlNamespaceManager(lobjNameTable)
          ''lobjNamespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
          ''Dim lobjParserContext As New XmlParserContext(lobjNameTable, lobjNamespaceManager, Nothing, XmlSpace.None)

          Dim lobjRegex As New Regex("<\?xml .*[^>]<[^ ]* ", RegexOptions.CultureInvariant Or RegexOptions.Compiled)
          Dim lobjMatch As Match = lobjRegex.Match(lpXmlString)

          Dim lobjOutputBuilder As New StringBuilder

          lobjOutputBuilder.Append(lpXmlString.Substring(0, lobjMatch.Length - 1))
          lobjOutputBuilder.AppendFormat("{0} ", lstrXSIRef)
          lobjOutputBuilder.Append(lpXmlString.Substring(lobjMatch.Length))

          Return lobjOutputBuilder.ToString

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

    End Class

    Partial Public Class Deserialize

      Shared Function XmlFile(ByVal lpFilePath As String, ByVal lpType As Type) As Object
        Dim lobjReader As XmlTextReader = Nothing
        Dim lobjObject As Object = Nothing

        ' ApplicationLogging.WriteLogEntry("Enter Deserialize::XmlFile", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
          lobjReader = New XmlTextReader(lpFilePath)
          lobjObject = lobjFormatter.Deserialize(lobjReader)
        Catch ex As InvalidOperationException
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          If ex.InnerException IsNot Nothing Then
            'Throw New InvalidOperationException(ex.InnerException.Message, ex)
            ' Throw a DeserializationException with the inner exception
            Throw New Exceptions.DeserializationException(ex.InnerException.Message, ex.InnerException)
          Else
            '  Re-throw the exception to the caller
            Throw
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New InvalidOperationException("Unable to deserialize xml file. '" & lpFilePath & "' [" & ex.Message & "]", ex)
        Finally
          If Not lobjReader Is Nothing Then lobjReader.Close()
          ' ApplicationLogging.WriteLogEntry("Exit Deserialize::XmlFile", TraceEventType.Verbose)
        End Try

        Return lobjObject

      End Function

      'Shared Function XmlStream(ByVal lpStream As IO.Stream, ByVal lpType As Type) As Object
      '  Dim lobjReader As XmlTextReader = Nothing
      '  Dim lobjObject As Object = Nothing

      '  ApplicationLogging.WriteLogEntry("Enter Deserialize::XmlFile", TraceEventType.Verbose)
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      '    lobjReader = New XmlTextReader(lpStream)
      '    lobjObject = lobjFormatter.Deserialize(lobjReader)
      '  Catch ex As InvalidOperationException
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    If ex.InnerException IsNot Nothing Then
      '      'Throw New InvalidOperationException(ex.InnerException.Message, ex)
      '      ' Throw a DeserializationException with the inner exception
      '      Throw New DeserializationException(ex.InnerException.Message, ex.InnerException)
      '    End If
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New DeserializationException(String.Format("Unable to deserialize xml file from stream. [{0}]", ex.Message), ex)
      '  Finally
      '    If Not lobjReader Is Nothing Then lobjReader.Close()
      '    ApplicationLogging.WriteLogEntry("Exit Deserialize::XmlFile", TraceEventType.Verbose)
      '  End Try

      '  Return lobjObject

      'End Function

      Shared Function XmlString(ByVal lpXml As String, ByVal lpType As Type) As Object
        Dim lobjReader As System.IO.StringReader = Nothing
        ApplicationLogging.WriteLogEntry("Enter Deserialize::XmlString", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
          lobjReader = New System.IO.StringReader(lpXml)
          Return lobjFormatter.Deserialize(lobjReader)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Dim lstrErrorDetail As String
          If ex.InnerException IsNot Nothing Then
            lstrErrorDetail = String.Format("{0}: {1}", ex.Message, ex.InnerException.Message)
          Else
            lstrErrorDetail = ex.Message
          End If

          ' Throw a DeserializationException 
          Throw New Exceptions.DeserializationException(String.Format("Unable to deserialize xml string '{0}'", lstrErrorDetail), ex)

        Finally
          If Not lobjReader Is Nothing Then lobjReader.Close()
          ApplicationLogging.WriteLogEntry("Exit Deserialize::XmlString", TraceEventType.Verbose)
        End Try
      End Function

      'Shared Function CustomFile(ByVal lpFilePath As String, ByVal lpFormatter As IFormatter) As Object
      '  Dim lobjReader As FileStream = Nothing
      '  ' ApplicationLogging.WriteLogEntry("Enter Deserialize::CustomFile", TraceEventType.Verbose)
      '  Try
      '    lobjReader = File.OpenRead(lpFilePath)
      '    Return lpFormatter.Deserialize(lobjReader)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  Finally
      '    If Not lobjReader Is Nothing Then lobjReader.Close()
      '    ' ApplicationLogging.WriteLogEntry("Exit Deserialize::CustomFile", TraceEventType.Verbose)
      '  End Try
      'End Function

      Public Shared Function ReadFileToMemory(ByVal lpFilePath As String) As MemoryStream
        Try

          ' Check to see if we have a valid file path
          Helper.VerifyFilePath(lpFilePath, True)

          Dim lintChunkSize As Integer = 1024
          Dim lobjBuffer(lintChunkSize - 1) As Byte
          Dim lintBytesRead As Integer = 0
          Dim lobjMemoryStream As New IO.MemoryStream()

          Using lobjFileStream As FileStream = File.OpenRead(lpFilePath)
            lobjMemoryStream.SetLength(lobjFileStream.Length)
            lintBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)

            While (lintBytesRead > 0)
              lobjMemoryStream.Write(lobjBuffer, 0, lintBytesRead)
              lintBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)
            End While

          End Using

          Return lobjMemoryStream

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Function

      Shared Function FromZippedStream(ByVal lpStream As Stream,
                                       ByVal lpZipFile As ZipFile,
                                       ByVal lpFormatter As XmlSerializer) As Object
        ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromZippedStream", TraceEventType.Verbose)
        Try
          If (lpStream.CanSeek) Then
            lpStream.Position = 0
          End If
          Dim lobjReader As New XmlZipStreamReader(lpStream, lpZipFile)
          Return lpFormatter.Deserialize(lobjReader)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New Exceptions.DeserializationException(ex.Message, ex)
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromZippedStream", TraceEventType.Verbose)
        End Try
      End Function

      Shared Function FromZippedStream(ByVal lpStream As Stream,
                                       ByVal lpZipFile As ZipFile,
                                       ByVal lpType As Type) As Object
        ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromZippedStream", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
          'Dim lobjFormatter As XmlStreamSerializer = New XmlStreamSerializer(lpStream, lpType)
          Return FromZippedStream(lpStream, lpZipFile, lobjFormatter)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Throw New Exceptions.DeserializationException(ex.Message, ex)
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromZippedStream", TraceEventType.Verbose)
        End Try
      End Function

      Shared Function FromStream(ByVal lpStream As Stream, ByVal lpFormatter As XmlSerializer) As Object
        ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromStream", TraceEventType.Verbose)
        Try
          If (lpStream.CanSeek) Then
            lpStream.Position = 0
          End If
          Dim lobjReader As New XmlStreamReader(lpStream)
          Return lpFormatter.Deserialize(lobjReader)
        Catch InvOpEx As InvalidOperationException
          If InvOpEx.Message = "There is an error in XML document (1, 1)." Then
            Dim lobjDeserializeException As New NonXmlDeserializationException("The stream is not an xml document.", InvOpEx)
            ApplicationLogging.LogException(lobjDeserializeException, Reflection.MethodBase.GetCurrentMethod, 62961)
            Throw lobjDeserializeException
          Else
            ApplicationLogging.LogException(InvOpEx, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller inside a Deserialization exception
            Throw New DeserializationException(InvOpEx.Message, InvOpEx)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller inside a Deserialization exception
          Throw New DeserializationException(ex.Message, ex)
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromStream", TraceEventType.Verbose)
        End Try
      End Function

      Shared Function FromStream(ByVal lpStream As Stream, ByVal lpType As Type) As Object
        ApplicationLogging.WriteLogEntry("Exit Deserialize::FromStream", TraceEventType.Verbose)
        Try
          Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
          'Dim lobjFormatter As XmlStreamSerializer = New XmlStreamSerializer(lpStream, lpType)
          Return FromStream(lpStream, lobjFormatter)
        Catch ex As Exception
          If TypeOf ex Is DeserializationException Then
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          Else
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            Throw New DeserializationException(ex.Message, ex)
          End If
        Finally
          ApplicationLogging.WriteLogEntry("Exit Deserialize::FromStream", TraceEventType.Verbose)
        End Try
      End Function

      'Shared Function FromBinaryStream(ByVal lpStream As Stream) As Object
      '  ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromBinaryStream", TraceEventType.Verbose)
      '  Try
      '    Dim lobjFormatter As New BinaryFormatter
      '    If (lpStream.CanSeek) Then
      '      lpStream.Position = 0
      '    End If
      '    Return lobjFormatter.Deserialize(lpStream)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  Finally
      '    ' ApplicationLogging.WriteLogEntry("Exit Deserialize::FromBinaryStream", TraceEventType.Verbose)
      '  End Try
      'End Function

      'Shared Function CustomArray(ByVal lpBuffer As Byte(), ByVal lpFormatter As IFormatter) As Object
      '  Dim lobjReader As MemoryStream = Nothing
      '  ' ApplicationLogging.WriteLogEntry("Exit Deserialize::CustomArray", TraceEventType.Verbose)
      '  Try
      '    lobjReader = New MemoryStream(lpBuffer)
      '    Return lpFormatter.Deserialize(lobjReader)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  Finally
      '    If Not lobjReader Is Nothing Then lobjReader.Close()
      '    ' ApplicationLogging.WriteLogEntry("Exit Deserialize::CustomArray", TraceEventType.Verbose)
      '  End Try
      'End Function

      ''Shared Function CustomArray(ByVal lpBuffer As Byte(), ByVal lpFormatter As XmlSerializer) As Object
      ''  Dim lobjReader As MemoryStream = Nothing
      ''  ApplicationLogging.WriteLogEntry("Exit Deserialize::CustomArray", TraceEventType.Verbose)
      ''  Try
      ''    lobjReader = New MemoryStream(lpBuffer)
      ''    Return lpFormatter.Deserialize(lobjReader)
      ''  Catch ex As Exception
      ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ''    Throw New DeserializationException(ex.Message, ex)
      ''  Finally
      ''    If Not lobjReader Is Nothing Then lobjReader.Close()
      ''    ApplicationLogging.WriteLogEntry("Exit Deserialize::CustomArray", TraceEventType.Verbose)
      ''  End Try
      ''End Function

      'Shared Function CustomString(ByVal lpText As String, ByVal lpFormatter As IFormatter) As Object
      '  Try
      '    Dim lobjBuffer As Byte() = Encoding.Default.GetBytes(lpText)
      '    Return CustomArray(lobjBuffer, lpFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      'Shared Function CustomString(ByVal lpText As String, ByVal lpFormatter As XmlSerializer) As Object
      '  Try
      '    Dim lobjBuffer As Byte() = Encoding.Default.GetBytes(lpText)
      '    Return CustomArray(lobjBuffer, lpFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      'Shared Function CustomString(ByVal lpText As String, ByVal lpFormatter As XmlSerializer, ByVal lpEncoding As Encoding) As Object
      '  Try
      '    Dim lobjBuffer As Byte() = lpEncoding.GetBytes(lpText)
      '    Return CustomArray(lobjBuffer, lpFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      ''Shared Function SoapArray(ByVal lpBuffer As Byte(), ByVal lpType As Type) As Object
      ''  Try
      ''    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      ''    Return CustomArray(lpBuffer, lobjFormatter)
      ''  Catch ex As Exception
      ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ''    Throw New DeserializationException(ex.Message, ex)
      ''  End Try
      ''End Function

      'Shared Function SoapFile(ByVal lpFilePath As String, ByVal lpType As Type) As Object
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      '    Return CustomFile(lpFilePath, lobjFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      'Shared Function SoapString(ByVal lpSoap As String, ByVal lpType As Type) As Object
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      '    Return CustomString(lpSoap, lobjFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      'Shared Function SoapString(ByVal lpSoap As String, ByVal lpType As Type, ByVal lpEncoding As Encoding) As Object
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      '    Return CustomString(lpSoap, lobjFormatter, lpEncoding)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

      ''Shared Function BinaryArray(ByVal lpBuffer As Byte(), ByVal lpType As Type) As Object
      ''  Try
      ''    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      ''    Return CustomArray(lpBuffer, lobjFormatter)
      ''  Catch ex As Exception
      ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ''    Throw New DeserializationException(ex.Message, ex)
      ''  End Try
      ''End Function

      'Shared Function BinaryFile(ByVal lpFilePath As String, ByVal lpType As Type) As Object
      '  Try
      '    Dim lobjFormatter As XmlSerializer = New XmlSerializer(lpType)
      '    Return CustomFile(lpFilePath, lobjFormatter)
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    Throw New Exceptions.DeserializationException(ex.Message, ex)
      '  End Try
      'End Function

    End Class

  End Class

End Namespace

