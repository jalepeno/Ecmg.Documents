'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  EnumDescriptionTypeConverter.vb
'   Description :  [type_description_here]
'   Created     :  6/17/2013 1:50:53 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection
Imports Documents.Utilities

#End Region

Namespace TypeConverters

  Public Class EnumDescriptionTypeConverter
    Inherits EnumConverter

#Region "Class Variables"
    Private mobjType As Type

#End Region

#Region "Public Properties"

#End Region

#Region "Constructors"

    Public Sub New(lpType As Type)
      MyBase.New(lpType)
      Try
        mobjType = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    '''    Converts the specified value object to an enumeration object. 
    ''' </summary>
    ''' <param name="context" type="System.ComponentModel.ITypeDescriptorContext">
    '''     <para>
    '''         An <see cref="ITypeDescriptorContext"/> that provides a format context.
    '''     </para>
    ''' </param>
    ''' <param name="culture" type="System.Globalization.CultureInfo">
    '''     <para>
    '''         An optional <see cref="CultureInfo"/>. If not supplied, the current culture is assumed.
    '''     </para>
    ''' </param>
    ''' <param name="value" type="Object">
    '''     <para>
    '''         The <see cref="Object"/> to convert.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     An <see cref="Object"/> that represents the converted <paramref name="value"/>.
    ''' </returns>
    ''' <exception cref="FormatException"><paramref name="value"/> is not a valid value for the target type. </exception>
    ''' <exception cref="NotSupportedException">The conversion cannot be performed. </exception>
    Public Overrides Function ConvertFrom(context As ITypeDescriptorContext, culture As CultureInfo, value As Object) As Object
      Try
        Dim lstrValueString As String = TryCast(value, String)
        If Not String.IsNullOrEmpty(lstrValueString) Then
          Return GetValue(mobjType, lstrValueString)
        End If
        Return MyBase.ConvertFrom(context, culture, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     Converts the given value object to the specified destination type.
    ''' </summary>
    ''' <param name="context" type="System.ComponentModel.ITypeDescriptorContext">
    '''     <para>
    '''        An <see cref="ITypeDescriptorContext"/> that provides a format context. 
    '''     </para>
    ''' </param>
    ''' <param name="culture" type="System.Globalization.CultureInfo">
    '''     <para>
    '''         An optional <see cref="CultureInfo"/>. If not supplied, the current culture is assumed.
    '''     </para>
    ''' </param>
    ''' <param name="value" type="Object">
    '''     <para>
    '''        The <see cref="Object"/> to convert. 
    '''     </para>
    ''' </param>
    ''' <param name="destinationType" type="System.Type">
    '''     <para>
    '''         The <see cref="Type"/> to convert the value to.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     An <see cref="Object"/> that represents the converted <paramref name="value"/>.
    ''' </returns>
    ''' <exception cref="ArgumentNullException"><paramref name="destinationType"/> is null. </exception>
    ''' <exception cref="ArgumentException"><paramref name="value"/> is not a valid value for the enumeration. </exception>
    ''' <exception cref="NotSupportedException">The conversion cannot be performed. </exception>
    Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, ByVal culture As CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object
      Try
        If destinationType Is Nothing Then
          Throw New ArgumentNullException("destinationType")
        End If

        If value IsNot Nothing AndAlso GetType(String) Is destinationType Then
          Return GetDescription(mobjType, value.ToString())
        End If

        Return MyBase.ConvertTo(context, culture, value, destinationType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     This method will get the "description" of the given enumeration field name for
    '''     the given type (set by using the DescriptionAttribute). If there is no
    '''     description then it will simply return the given field name.
    ''' </summary>
    ''' <param name="lpType" type="System.Type">
    '''     <para>
    '''         The enumeration type to get the description for.
    '''     </para>
    ''' </param>
    ''' <param name="lpFieldName" type="String">
    '''     <para>
    '''         The enumeration fieldName to get the description for.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     The description of the given enumeration field name for the given type, the given
    '''     field name, or string.Empty if all else fails.
    ''' </returns>
    Public Shared Function GetDescription(ByVal lpType As Type, ByVal lpFieldName As String) As String
      Try
        If lpType IsNot Nothing AndAlso Not String.IsNullOrEmpty(lpFieldName) Then
          Dim lobjFieldInfo As FieldInfo = lpType.GetField(lpFieldName)
          If lobjFieldInfo IsNot Nothing Then
            Dim lobjAttributeArray() As DescriptionAttribute
            lobjAttributeArray = TryCast(lobjFieldInfo.GetCustomAttributes(GetType(DescriptionAttribute), False),
              DescriptionAttribute())
            If lobjAttributeArray IsNot Nothing AndAlso lobjAttributeArray.Length <> 0 Then
              Return lobjAttributeArray(0).Description
            End If
          End If
          Return lpFieldName
        End If
        Return String.Empty
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     Gets the value of an Enum, based on it's DescriptionAttribute or named value.
    ''' </summary>
    ''' <param name="lpType" type="System.Type">
    '''     <para>
    '''         The enumeration type to get the value for.
    '''     </para>
    ''' </param>
    ''' <param name="lpDescription" type="String">
    '''     <para>
    '''         The description or name of the element.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     The value, or the passed in description, if it was not found.
    ''' </returns>
    Public Shared Function GetValue(ByVal lpType As Type, ByVal lpDescription As String) As Object
      Try
        Dim lobjFieldInfoArray() As FieldInfo = lpType.GetFields
        For Each lobjFieldInfo As FieldInfo In lobjFieldInfoArray
          Dim lobjAttributeArray() As DescriptionAttribute
          lobjAttributeArray = TryCast(lobjFieldInfo.GetCustomAttributes(GetType(DescriptionAttribute), False),
            DescriptionAttribute())
          If lobjAttributeArray IsNot Nothing AndAlso lobjAttributeArray.Length <> 0 Then
            If lobjAttributeArray(0).Description = lpDescription Then
              Return lobjFieldInfo.GetValue(lobjFieldInfo.Name)
            End If
          End If

          If lobjFieldInfo.Name = lpDescription Then
            Return lobjFieldInfo.GetValue(lobjFieldInfo.Name)
          End If
        Next

        Return lpDescription

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetSortedDescriptions(ByVal lpType As Type) As IList(Of String)
      Try
        Dim lobjDescriptions As New List(Of String)

        If lpType IsNot Nothing Then
          ' Dim lobjFieldInfo As FieldInfo = lpType.GetField(lpFieldName)
          Dim lobjFieldInfo() As FieldInfo = lpType.GetFields()

          'If lobjFieldInfo IsNot Nothing Then
          '  Dim lobjAttributeArray() As DescriptionAttribute
          '  lobjAttributeArray = TryCast(lobjFieldInfo.GetCustomAttributes(GetType(DescriptionAttribute), False),  _
          '    DescriptionAttribute())
          '  If lobjAttributeArray IsNot Nothing AndAlso lobjAttributeArray.Length <> 0 Then
          '    Return lobjAttributeArray(0).Description
          '  End If
          'End If
          'Return lpFieldName
          If lobjFieldInfo IsNot Nothing Then
            Dim lobjAttributeArray() As DescriptionAttribute
            For Each lobjField As FieldInfo In lobjFieldInfo
              lobjAttributeArray = TryCast(lobjField.GetCustomAttributes(GetType(DescriptionAttribute), False),
              DescriptionAttribute())
              If lobjAttributeArray IsNot Nothing AndAlso lobjAttributeArray.Length <> 0 Then
                lobjDescriptions.Add(lobjAttributeArray.First.Description)
              End If
            Next
          End If

        End If

        lobjDescriptions.Sort()

        Return lobjDescriptions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

#End Region

  End Class

End Namespace