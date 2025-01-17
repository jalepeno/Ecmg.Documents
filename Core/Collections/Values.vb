'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Contains the values of a multi-valued property.</summary>
  Partial Public Class Values
    Inherits CCollection(Of Object)
    Implements IEquatable(Of Values)

#Region "Class Enumerations"

    Public Enum ValueIndexEnum
      First = -100
      Last = -101
    End Enum

#End Region

#Region "Class Variables"

    Private mobjParentProperty As ECMProperty = Nothing

#End Region

#Region "Overloads"

    Public Overloads Function Contains(ByVal lpValue As Core.Value) As Boolean
      Try
        For Each lobjValue As Object In Me
          If lobjValue.ToString = lpValue.ToString Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the collection of values contains a value with the specified string value.
    ''' </summary>
    ''' <param name="lpValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Contains(ByVal lpValue As String) As Boolean
      Try
        For Each lobjValue As Object In Me
          If lobjValue = lpValue Then
            Return True
          End If
        Next
        For Each lobjValue As String In Me
          If lobjValue = lpValue Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub Add(ByVal item As Object)
      Try
        If TypeOf (item) Is Value Then
          Add(item, True)
        ElseIf TypeOf (item) Is String Then
          AddString(item)
        ElseIf TypeOf (item) Is Boolean Then
          AddBoolean(item)
        ElseIf TypeOf (item) Is DateTime Then
          AddDate(item)
        ElseIf TypeOf (item) Is Double Then
          AddDouble(item)
        ElseIf TypeOf (item) Is Guid Then
          AddString(item.ToString)
        ElseIf TypeOf (item) Is Long Then
          AddLong(item)
        ElseIf TypeOf (item) Is Integer Then
          AddInteger(item)
        ElseIf TypeOf (item) Is Object Then
          AddObject(item)
        Else
          AddString(item)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As Value)
      Try
        Add(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddString(ByVal item As String)
      Try
        AddString(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddSingle(ByVal item As Single)
      Try
        AddSingle(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddDouble(ByVal item As Double)
      Try
        AddDouble(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddInteger(ByVal item As Integer)
      Try
        AddInteger(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddLong(ByVal item As Long)
      Try
        AddLong(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddBoolean(ByVal item As Boolean)
      Try
        AddBoolean(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddDate(ByVal item As Date)
      Try
        AddDate(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub AddObject(ByVal item As Object)
      Try
        AddObject(item, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub Add(ByVal item As Value, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item) Then
          Throw New Exceptions.ValueExistsException(item.ToString,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item.Value))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As String, ByVal allowDuplicates As Boolean)
      Try
        AddString(item, allowDuplicates)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddString(ByVal item As String, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddSingle(ByVal item As Single, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddDouble(ByVal item As Double, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddInteger(ByVal item As Integer, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddLong(ByVal item As Long, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddBoolean(ByVal item As Boolean, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddDate(ByVal item As Date, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a value to the collection
    ''' </summary>
    ''' <param name="item">The item to be added</param>
    ''' <param name="allowDuplicates">Specifies whether or not 
    ''' to allow duplicate items to be added to the collection</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If allowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Shadows Sub AddObject(ByVal item As Object, ByVal allowDuplicates As Boolean)
      Try
        If allowDuplicates = False AndAlso Contains(item.ToString) Then
          Throw New Exceptions.ValueExistsException(item,
                                                    String.Format("Value '{0}' not added because it already exists.",
                                                                  item))
          'ApplicationLogging.WriteLogEntry(String.Format("Value '{0}' not added because it already exists.", item.Value), TraceEventType.Warning, 45326)
        Else
          'MyBase.Add(New Value(item))
          MyBase.Add(item)
        End If
      Catch ValueExistsEx As Exceptions.ValueExistsException
        ' Don't log it here, just pass it on.
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    Default Overrides Property Item(ByVal index As Integer) As Object
      Get
        Try
          Dim lobjValue As Object = MyBase.Item(index)
          If TypeOf lobjValue Is Value Then
            Return CType(lobjValue, Value).Value
          Else
            Return lobjValue
          End If
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Object)
        Try
          MyBase.Item(index) = value
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property ParentProperty As ECMProperty
      Get
        Return mobjParentProperty
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpParentProperty As ECMProperty)
      MyBase.New()
      mobjParentProperty = lpParentProperty
    End Sub

#End Region

#Region "Public Operator Overloads"

    Public Shared Operator =(ByVal v1 As Values, ByVal v2 As Values) As Boolean
      ' Compare all the parts of the criterion
      'Return c1.Cardinality = c2.Cardinality AndAlso _
      '       c1.Name = c2.Name AndAlso _
      '       c1.Operator = c2.Operator AndAlso _
      '       c1.PropertyName = c2.PropertyName AndAlso _
      '       c1.PropertyScope = c2.PropertyScope AndAlso _
      '       c1.SetEvaluation = c2.SetEvaluation AndAlso _
      '       c1.Value = c2.Value AndAlso c1.Values.ToString = c2.Values.ToString
      Try

        If v1 IsNot Nothing AndAlso v2 Is Nothing Then
          Return False
        End If

        If v1 Is Nothing AndAlso v2 IsNot Nothing Then
          Return False
        End If

        If v1.Count <> v2.Count Then
          Return False
        End If

        For lintValueCounter As Integer = 0 To v1.Count - 1

          If v1.Item(lintValueCounter).GetType.IsEnum AndAlso Not v2.Item(lintValueCounter).GetType.IsEnum Then
            If v1.Item(lintValueCounter).ToString = v2.Item(lintValueCounter).ToString() Then
              Return True
            Else
              Return False
            End If
          End If

#If CTSDOTNET = 1 Then
          If IsNumeric(v1.Item(lintValueCounter)) AndAlso IsNumeric(v2.Item(lintValueCounter)) Then
            If v1.Item(lintValueCounter) - v2.Item(lintValueCounter) = 0 Then
              Return True
            Else
              Return False
            End If
          End If
#End If

          If (IsDate(v1.Item(lintValueCounter)) AndAlso IsDate(v2.Item(lintValueCounter))) Then
            Dim lobjTimeDifference As TimeSpan = DirectCast(v1.Item(lintValueCounter), DateTime) -
                                                 DirectCast(v1.Item(lintValueCounter), DateTime)
            If lobjTimeDifference.TotalSeconds = 0 Then
              Return True
            Else
              Return False
            End If
          End If

          If v1.Item(lintValueCounter).Equals(v2.Item(lintValueCounter)) = False Then
            Return False
          End If
        Next

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator <>(ByVal v1 As Values, ByVal v2 As Values) As Boolean
      Try
        Return Not (v1 = v2)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Overloads Function Equals(ByVal v As Values) As Boolean Implements System.IEquatable(Of Values).Equals
      Try
        If v Is Nothing Then
          Return False
        Else
          Return v = Me
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Shared Function Equals(ByVal v1 As Values, ByVal v2 As Values) As Boolean
      Try
        Return (v1 = v2)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

    Public Function GetFirstValue() As Object
      Try
        If Me.Count > 0 Then
          Return Me.Item(0)
        Else
          Throw New InvalidOperationException("The value collection is empty.")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetLastValue() As Object
      Try
        Dim lintTotalCount As Integer = Me.Count
        If lintTotalCount > 0 Then
          Return Me.Item(lintTotalCount - 1)
        Else
          Throw New InvalidOperationException("The value collection is empty.")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToDelimitedString(ByVal lpDelimiter As String) As String
      Try

        If Me.Count > 0 Then
          Dim lobjReturnStringBuilder As New Text.StringBuilder

          For Each lobjValue As Object In Me
            If ((lobjValue IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(lobjValue.ToString()))) Then
              lobjReturnStringBuilder.AppendFormat("{0}{1}", lobjValue.ToString(), lpDelimiter)
            End If
          Next

          If lobjReturnStringBuilder.Length - lpDelimiter.Length >= 0 Then
            lobjReturnStringBuilder = lobjReturnStringBuilder.Remove(lobjReturnStringBuilder.Length - lpDelimiter.Length,
                                                                     lpDelimiter.Length)
          End If

          Return lobjReturnStringBuilder.ToString()

        Else
          Return String.Empty
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToStringArray() As String()
      Try
        Dim lobjStringArray As New List(Of String)
        If ParentProperty IsNot Nothing Then
          Select Case ParentProperty.Type
            Case PropertyType.ecmBinary, PropertyType.ecmObject, PropertyType.ecmUndefined
              Throw New Exceptions.InvalidPropertyTypeException(
                String.Format("Unable to convert values to string array {0} type not supported.",
                              ParentProperty.Type.ToString.Substring(3)), ParentProperty)
            Case Else
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjStringArray.Add(lobjValue.Value)
                Else
                  lobjStringArray.Add(lobjValue)
                End If
              Next
          End Select

          Return lobjStringArray.ToArray

        Else
          Return MyBase.ToArray
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function ToArray() As Array
      Try

        If ParentProperty IsNot Nothing Then
          Select Case ParentProperty.Type
            Case PropertyType.ecmBinary, PropertyType.ecmObject, PropertyType.ecmUndefined
              Throw New Exceptions.InvalidPropertyTypeException(
                String.Format("Unable to convert values to string array {0} type not supported.",
                              ParentProperty.Type.ToString.Substring(3)), ParentProperty)

            Case PropertyType.ecmString
              Dim lobjArray As New List(Of String)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case PropertyType.ecmBoolean
              Dim lobjArray As New List(Of Boolean)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case PropertyType.ecmDate
              Dim lobjArray As New List(Of DateTime)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case PropertyType.ecmDouble
              Dim lobjArray As New List(Of Double)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case PropertyType.ecmGuid
              Dim lobjArray As New List(Of String)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case PropertyType.ecmLong
              Dim lobjArray As New List(Of Long)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjArray.Add(lobjValue.Value)
                Else
                  lobjArray.Add(lobjValue)
                End If
              Next
              Return lobjArray.ToArray
            Case Else
              Dim lobjStringArray As New List(Of String)
              For Each lobjValue As Object In Me
                If TypeOf (lobjValue) Is Value Then
                  lobjStringArray.Add(lobjValue.Value)
                Else
                  lobjStringArray.Add(lobjValue)
                End If
              Next
              Return lobjStringArray.ToArray
          End Select
        Else
          Return MyBase.ToArray
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Sub SetParentProperty(ByVal lpParentProperty As ECMProperty)
      Try
        mobjParentProperty = lpParentProperty
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region
  End Class
End Namespace