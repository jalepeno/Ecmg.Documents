'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Globalization
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json



#End Region

Namespace Transformations
  ''' <summary>An action to be performed in a transformation operation.</summary>
  ''' <remarks>The base class from which all transformation actions inherit.</remarks>
  <XmlInclude(GetType(PropertyAction)),
     XmlInclude(GetType(ContentFileHasExtensionTestAction)),
     XmlInclude(GetType(RunTransformationAction)),
     XmlInclude(GetType(ExitTransformationAction)),
     XmlInclude(GetType(RemoveVersionsWithoutContentAction)),
     XmlInclude(GetType(DocumentClassTestAction)),
     XmlInclude(GetType(TargetClassAction)),
     XmlInclude(GetType(TransformAllRelatedDocumentsAction)),
     XmlInclude(GetType(ClearAllFolderPaths)),
     XmlInclude(GetType(AddLiteralFolderPath)),
     XmlInclude(GetType(ChangeFolderPath)),
     XmlInclude(GetType(ReplaceFolderPath)),
     XmlInclude(GetType(RemoveFolderLevel))>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class Action
    Implements IDescription
    Implements IActionable
    Implements IActionableInformation
    Implements IActionItem
    Implements ITransformationAction

#Region "Class Constants"

    Friend Const CATEGORY_CONFIG As String = "Configuration"

#End Region

#Region "Class Variables"

    Private mstrId As String
    Private menuActionType As ActionType
    Protected mstrName As String = String.Empty
    Private mobjTransformation As New Transformation
    Protected mstrDescription As String = String.Empty
    Private WithEvents mobjParameters As IParameters = New Parameters

    Protected menuResult As Result = Result.NotProcessed
    Private mstrInstanceId As String = String.Empty
    Private mobjLocale As CultureInfo = CultureInfo.CurrentCulture

#End Region

#Region "Public Properties"

    <XmlAttribute("id")>
    Public Property Id As String
      Get
        Try
          If String.IsNullOrEmpty(mstrId) Then
            mstrId = Guid.NewGuid.ToString()
          End If
          Return mstrId
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          If Helper.IsDeserializationBasedCall OrElse
            Helper.CallStackContainsMethodName("AssignObjectProperties") OrElse
            Helper.CallStackContainsMethodName("Clone") Then
            mstrId = value
          Else
            Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable ReadOnly Property ActionName As String Implements IActionableBase.Name, IActionableInformation.Name, IActionItem.Name
      Get
        Try
          Return Name
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlAttribute()>
    Public Overridable Property Name() As String Implements IDescription.Name, INamedItem.Name
      Get
        If String.IsNullOrEmpty(mstrName) Then
          mstrName = Me.GetType.Name
        End If
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    <XmlAttribute()>
    Public Property Description() As String Implements IDescription.Description
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Overridable ReadOnly Property DisplayName As String Implements ITransformationAction.DisplayName, IActionable.DisplayName, IActionableInformation.DisplayName, IActionItem.DisplayName
      Get
        Try
          Return Helper.CreateDisplayName(Me.Name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <Category(CATEGORY_CONFIG)>
    <XmlIgnore()>
    Public Property Parameters As IParameters Implements ITransformationAction.Parameters, IActionable.Parameters, IActionItem.Parameters
      Get
        Try
          If mobjParameters.Count = 0 Then
            mobjParameters = GetDefaultParameters()
          End If
          Return mobjParameters
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IParameters)
        Try
          mobjParameters = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property ProcessedMessage As String Implements IActionableBase.ProcessedMessage

    Public ReadOnly Property Result As Result Implements IActionableBase.Result
      Get
        Return menuResult
      End Get
    End Property

    Public ReadOnly Property Type() As ActionType
      Get
        Return menuActionType
      End Get
    End Property

    <XmlIgnore()>
    Public Property Transformation() As Transformation
      Get
        Return mobjTransformation
      End Get
      Set(ByVal value As Transformation)
        mobjTransformation = value
      End Set
    End Property

    Public ReadOnly Property ResultDetail As IActionableResult Implements IActionable.ResultDetail
      Get
        Try
          ' Throw New NotImplementedException
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public Property ActionableParent As IActionableBase Implements IActionableBase.ActionableParent

    Public ReadOnly Property ActionDescription As String Implements IActionableBase.Description, IActionableInformation.Description, IActionItem.Description
      Get
        Try
          Return Me.Description
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property DocumentId As String Implements IActionable.DocumentId

    <XmlIgnore()>
    Public Property Host As Object Implements IActionable.Host

    Public ReadOnly Property InstanceId As String Implements IActionable.InstanceId
      Get
        Return mstrInstanceId
      End Get
    End Property

    Public ReadOnly Property IsDisposed As Boolean Implements IActionable.IsDisposed
      Get
        Return disposedValue
      End Get
    End Property

    Public ReadOnly Property Locale As Globalization.CultureInfo Implements IActionable.Locale
      Get
        Return mobjLocale
      End Get
    End Property

    <XmlIgnore()>
    Public Property ShouldExecute As Boolean Implements IActionable.ShouldExecute

    <XmlIgnore()>
    Public Property Tag As Object Implements IActionable.Tag

    Public ReadOnly Property CompanyName As String Implements IActionableInformation.CompanyName
      Get
        Return ConstantValues.CompanyName
      End Get
    End Property

    Public Property ExtensionPath As String Implements IActionableInformation.ExtensionPath

    Public ReadOnly Property IsExtension As Boolean Implements IActionableInformation.IsExtension
      Get
        Return False
      End Get
    End Property

    Public ReadOnly Property ProductName As String Implements IActionableInformation.ProductName
      Get
        Return ConstantValues.ProductName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        mstrName = Me.GetType.Name
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpActionType As ActionType)
      Me.New(lpActionType, "")
    End Sub

    Public Sub New(ByVal lpActionType As ActionType, ByVal lpName As String)
      menuActionType = lpActionType
      mstrName = lpName
    End Sub

#End Region

#Region "Public Methods"

    Public Sub SetInstanceId(lpInstanceId As String) Implements IActionable.SetInstanceId
      Try
        mstrInstanceId = lpInstanceId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Execute() As ActionResult Implements IActionable.Execute
      Try
        Return Execute(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public MustOverride Function Execute(ByRef lpErrorMessage As String) As ActionResult Implements IActionable.Execute

    Public Function OnExecute() As Result Implements IActionable.OnExecute
      Try
        Dim lobjActionResult As ActionResult = Execute()
        Return lobjActionResult.ToResultCode()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SetDescription(lpDescription As String) Implements IActionable.SetDescription, IActionItem.SetDescription
      Try
        mstrDescription = lpDescription
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub SetResult(lpResult As Result) Implements IActionable.SetResult
      Try
        menuResult = lpResult
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public MustOverride Sub InitializeFromParameters() Implements ITransformationAction.InitializeFromParameters

    Public Shared Function CreateFromJsonReader(reader As JsonReader) As Action
      Try
        'Return JsonConvert.DeserializeObject(reader, GetType(IOperation), New OperationJsonConverter())
        Dim lobjConverter As New ActionJsonConverter()
        Return lobjConverter.ReadJson(reader, Nothing, Nothing, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ToActionItem() As IActionItem
      Try
        Return New ActionItem(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function ToString() As String Implements IActionable.ToString
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToJson() As String
      Try
        Return WriteJsonString()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlString() As String Implements IActionable.ToXmlElementString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToXmlItemString() As String Implements IActionItem.ToXmlElementString
      Try
        Return Serializer.Serialize.XmlElementString(Me.ToActionItem())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
      Try
        Return TransformationActionFactory.Create(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend MustOverride Sub InitializeParameterValues() _
      Implements IActionable.InitializeParameterValues,
      ITransformationAction.InitializeParameterValues

    Protected Overridable Function GetCurrentParameters() As IParameters
      Try
        Return GetDefaultParameters()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetDefaultParameters() As IParameters
      Try
        Return New Parameters
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetBooleanParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean Implements IActionable.GetBooleanParameterValue, IActionItem.GetBooleanParameterValue
      Try
        Return ActionItem.GetBooleanParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetEnumParameterValue(ByVal lpParameterName As String, ByVal lpEnumType As Type, ByVal lpDefaultValue As Object) As [Enum] Implements IActionable.GetEnumParameterValue, IActionItem.GetEnumParameterValue
      Try
        Return ActionItem.GetEnumParameterValue(Me, lpParameterName, lpEnumType, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetStringParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As String Implements IActionable.GetStringParameterValue, IActionItem.GetStringParameterValue
      Try
        Return ActionItem.GetStringParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object Implements IActionable.GetParameterValue, IActionItem.GetParameterValue
      Try
        Return ActionItem.GetParameterValue(Me, lpParameterName, lpDefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Function WriteJsonString() As String
      Try
        Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.None, New ActionJsonConverter())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Friend Function DebuggerIdentifier() As String
      Try
        If String.IsNullOrEmpty(Name) Then
          Return Me.GetType.Name
        Else
          If String.Compare(Me.GetType.Name, Name) = 0 Then
            Return Name
          Else
            Return String.Format("{0}: {1}", Me.GetType.Name, Name)
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

    <XmlIgnore()>
    Public Property ITransformation As ITransformation Implements ITransformationAction.Transformation
      Get
        Try
          Return Me.Transformation
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As ITransformation)
        Try
          Me.Transformation = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

  End Class

End Namespace