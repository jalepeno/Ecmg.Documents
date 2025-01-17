'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core.ChoiceLists

  ''' <summary>
  ''' Represents a list of values or 'Controlled Vocabulary List'
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceList
    Inherits RepositoryObject
    'Implements IChoiceList
    Implements IEquatable(Of ChoiceList)
    Implements IStreamSerialize

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const CHOICE_LIST_FILE_EXTENSION As String = "cvl"

#End Region

#Region "Class Variables"

    Private menuChoiceType As ChoiceType = ChoiceType.ChoiceString
    Private mobjChoiceValues As New ChoiceValues

#End Region

#Region "Public Properties"

    Public Overrides Property Name() As String
      Get
        Return MyBase.Name
      End Get
      Set(ByVal value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Property Type() As ChoiceType
      Get
        Return menuChoiceType
      End Get
      Set(ByVal value As ChoiceType)
        menuChoiceType = value
      End Set
    End Property

    Public Property ChoiceValues() As ChoiceValues 'Implements IChoiceList.ChoiceValues
      Get
        Return mobjChoiceValues
      End Get
      Set(ByVal value As ChoiceValues)
        mobjChoiceValues = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    'Public Sub New(ByVal Name As String)
    '  ' We will set the Name and the DisplayName to the same value to start.
    '  MyBase.New(Name)
    'End Sub



    ''' <summary>
    ''' Constructs a new choice list object using the specified stream.
    ''' </summary>
    ''' <param name="lpStream">An IO.Stream object derived from a choice list file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpStream As IO.Stream)
      Try
        Dim lobjChoiceList As ChoiceList = DeSerialize(lpStream)
        Helper.AssignObjectProperties(lobjChoiceList, Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ChoiceList::New('{0}')", lpStream))
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    'Public Sub LoadFromFile(ByVal lpXMLFilePath As String)
    '  Try
    '    Dim lobjXmlDocument As New Xml.XmlDocument
    '    lobjXmlDocument.Load(lpXMLFilePath)

    '    LoadFromXmlDocument(lobjXmlDocument)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Helper.DumpException(ex)
    '  End Try
    'End Sub

    'Public Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
    '  Try
    '    Dim lobjChoiceList As ChoiceList = DeSerialize(lpXML)
    '    With Me
    '      .Id = lobjChoiceList.Id
    '      .Name = lobjChoiceList.Name
    '      .DescriptiveText = lobjChoiceList.DescriptiveText
    '      .DisplayName = lobjChoiceList.DisplayName
    '      .ChoiceValues = lobjChoiceList.ChoiceValues
    '    End With

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Helper.DumpException(ex)
    '  End Try
    'End Sub

    Public Sub Sort()
      Try
        ChoiceValues.Sort()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Compares the current choicelist to the specified one.
    ''' </summary>
    ''' <param name="choiceList">The choice list to compare to.</param>
    ''' <returns></returns>
    ''' <remarks>The id is ignored in the comparison.</remarks>
    Public Overloads Function Equals(ByVal choiceList As ChoiceList) As Boolean Implements IEquatable(Of ChoiceList).Equals
      Try

        If choiceList Is Nothing Then
          Return False
        End If

        With choiceList

          If .Name <> Name Then
            Return False
          End If

          If .DisplayName <> DisplayName Then
            Return False
          End If

          If .Properties.Count <> Properties.Count Then
            Return False
          End If

          For Each lobjProperty As ECMProperty In .Properties

            If Properties.Contains(lobjProperty.Name) = False Then
              Return False
            End If

            If Properties(lobjProperty.Name).Equals(lobjProperty) = False Then
              Return False
            End If
          Next

          For Each lobjChoiceValue As ChoiceValue In .ChoiceValues
            If ChoiceValues.Contains(lobjChoiceValue.Name) = False Then
              Return False
            End If

            ' TODO: Implement a direct choice value comparison as well
          Next

        End With

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' If the Name is not empty, returns the name, 
    ''' otherwise returns as the base object would.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>To return an XML string representation 
    ''' of this object, call ToXmlString.</remarks>
    Public Overrides Function ToString() As String
      Try
        If Name.Length > 0 Then
          Return Me.Name
        Else
          Return MyBase.ToString
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Function DeSerialize(ByVal lpFilePath As String) As Object
    '  Try
    '    Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

#Region "IStreamSerialize Implementation"

    Public Function DeSerialize(ByVal lpStream As System.IO.Stream) As Object Implements IStreamSerialize.DeSerialize
      Try
        Return Serializer.Deserialize.FromStream(lpStream, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    Public Function SerializeToStream() As System.IO.Stream Implements IStreamSerialize.SerializeToStream
      Try
        Return Serializer.Serialize.ToStream(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace