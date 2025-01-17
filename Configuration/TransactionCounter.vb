'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  TransactionCounter.vb
'   Description :  [type_description_here]
'   Created     :  8/29/2013 3:04:38 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Data
Imports System.Environment
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Timers
Imports Documents.Configuration
Imports Documents.Utilities
Imports Microsoft.Data.SqlClient


#End Region

Namespace Core
  ''' <exclude />
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TransactionCounter
    Implements IDisposable
    Implements INotifyPropertyChanged

    ' Used to manage the cumulative transaction counts on a given machine.

#Region "Class Constants"

    Private Const PARENT_FOLDER_NAME As String = "4f9c12ee-0487-4f5c-85a6-4de6b29e17b8"
    Protected Const TRANSACTION_COUNTER_FILE_NAME As String = "ptc"
    Private Const KEY As String = "ffQZ75KZgmqC0/to4Hb2qv3nMAx50NqUuQQLiAACmDI="
    Private Const COMMAND_TIMEOUT As Integer = 30

#End Region

#Region "Class Variables"

    Private Shared mobjInstance As TransactionCounter
    Private Shared mintReferenceCount As Integer
    Private Shared mstrTransactionCounterFilePath As String = String.Empty
    Private Shared mstrCatalogConnectionString As String = String.Empty
    Private Shared mlngCurrentCount As Long
    Private Shared mlngCurrentCounter As Long
    Private Shared ReadOnly mobjSymmetric As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
    Private Shared mobjEncryptedData As New Encryption.Data
    Private Shared ReadOnly mobjKey As New Encryption.Data()
    Protected Shared mstrFileName As String = TRANSACTION_COUNTER_FILE_NAME
    Private Shared mstrNodeName As String
    Private mblnIsDirty As Boolean = False

    Private Shared ReadOnly mobjTimer As New Timer(10000)

    Private Shared mblnCounterFileAccessAllowed As Boolean = True

    ' Private mobjCallback As New TimerCallback(AddressOf ProcessTimerEvent)

#End Region

#Region "Public Events"

    ''' <summary>Fired when the counter is incremented.</summary>
    Public Event CountIncremented(currentCount As Long)

    ''' <summary>Notifies of property changes.</summary>
    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) _
      Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Public Properties"

#Region "Singleton Support"

    ''' <summary>Returns the current instance of the TransactionCounter object in support of the Singleton object pattern.</summary>
    Public Shared Function Instance() As TransactionCounter

      Try

        If mobjInstance Is Nothing Then
          mobjInstance = New TransactionCounter
          mstrTransactionCounterFilePath = GetPrimaryFilePath()
          mstrCatalogConnectionString = ConnectionSettings.Instance.ProjectCatalogConnectionString

          mstrNodeName = Environment.MachineName

          mobjKey.Base64 = KEY
          'mlngCurrentCount = GetCurrentCountFromFile()

          ' For the log file
          Dim lintCurrentDBCount As Long = GetCurrentCountFromDB()
          Dim lintCurrentFileCount As Long = GetCurrentCountFromFile()
          ApplicationLogging.WriteLogEntry(String.Format("Current transaction count: (File={0:n0}, Db={1:n0})", lintCurrentFileCount, lintCurrentDBCount), MethodBase.GetCurrentMethod(), TraceEventType.Information, 12345)

          mlngCurrentCount = GetCurrentCount()
          mlngCurrentCounter = 0
          AddHandler mobjTimer.Elapsed, AddressOf mobjInstance.ProcessTimerEvent
          mobjTimer.Enabled = True
        End If

        mintReferenceCount += 1
        Return mobjInstance

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

#Region "Private Properties"

    ''' <summary>Gets the current counter file path.</summary>
    Private Shared ReadOnly Property FilePath As String
      Get

        Try

          If String.IsNullOrEmpty(mstrTransactionCounterFilePath) Then
            mstrTransactionCounterFilePath = GetPrimaryFilePath()
          End If

          Return mstrTransactionCounterFilePath

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Protected Shared ReadOnly Property FileName As String
      Get
        Try
          Return mstrFileName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      'Set(value As String)
      '  Try
      '    mstrFileName = value
      '  Catch ex As Exception
      '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '    ' Re-throw the exception to the caller
      '    Throw
      '  End Try
      'End Set
    End Property

    Public Shared ReadOnly Property CounterFileAccessAllowed As Boolean
      Get
        Try
          Return mblnCounterFileAccessAllowed
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Protected Sub New()
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>Gets the current counter value.</summary>
    Public Shared ReadOnly Property CurrentCount As Long
      Get

        Try
          Return mlngCurrentCount

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>Gets the last saved counter value.</summary>
    Public Shared ReadOnly Property SavedCount As Long
      Get

        Try
          Return GetCurrentCount()

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Friend Methods"

    ''' <summary>Sets the current count.</summary>
    Friend Sub SetCurrentCount(lpCurrentCount As Long)

      Try

        If lpCurrentCount <> mlngCurrentCount Then
          mlngCurrentCount = lpCurrentCount
          mblnIsDirty = True
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>Increments the current count by one.</summary>
    Public Sub Increment()

      Try
        mlngCurrentCount += 1
        mlngCurrentCounter += 1
        mblnIsDirty = True
        RaiseEvent CountIncremented(mlngCurrentCount)
        OnPropertyChanged("CurrentCount")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub SaveCurrentCount()
      Try
        'SaveCurrentCountToDB()
        TransactionCounter.SaveNewItemsToDB()
        If CounterFileAccessAllowed Then
          SaveCurrentCountToFile()
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub SaveCurrentCountToDB()
      Try
        Dim lobjSQLBuilder As New StringBuilder
        lobjSQLBuilder.AppendFormat("UPDATE tblNodes SET ItemsProcessed = {0} WHERE NodeAddress = '{1}'",
                                    mlngCurrentCount, mstrNodeName)
        Dim lintRecordsAffected As Integer = TransactionCounter.ExecuteNonQuery(lobjSQLBuilder.ToString(), mstrCatalogConnectionString)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub SaveNewItemsToDB()
      Try

        If String.IsNullOrEmpty(mstrCatalogConnectionString) Then
          Exit Sub
        End If

        Using lobjConnection As New SqlConnection(mstrCatalogConnectionString)
          Using lobjCommand As New SqlCommand("usp_update_node_counter", lobjConnection)
            lobjCommand.CommandType = CommandType.StoredProcedure
            lobjCommand.CommandTimeout = COMMAND_TIMEOUT

            Dim lobjProjectIdParameter As New SqlParameter("@nodeAddress", SqlDbType.NVarChar, 255) With {
              .Value = mstrNodeName
            }
            lobjCommand.Parameters.Add(lobjProjectIdParameter)

            Dim lobjJobIdParameter As New SqlParameter("@additionalItemCount", SqlDbType.BigInt) With {
              .Value = mlngCurrentCounter
            }
            lobjCommand.Parameters.Add(lobjJobIdParameter)

            Dim lobjReturnParameter As New SqlParameter("@returnvalue", SqlDbType.Int) With {
              .Direction = ParameterDirection.ReturnValue
            }
            lobjCommand.Parameters.Add(lobjReturnParameter)

            Helper.HandleConnection(lobjConnection)
            Dim lintReturnValue As Integer
            lobjCommand.ExecuteNonQuery()
            lintReturnValue = lobjReturnParameter.Value
            'If lintReturnValue = -100 Then
            '  Throw New Exception(String.Format("Failed to save batch '{0}'.", lpBatch.Id))
            'End If

            Select Case lintReturnValue
              Case -100
                Throw New InvalidOperationException("Null node address recieved by usp_update_node_counter")
              Case -200
                Throw _
                  New InvalidOperationException(String.Format("Node address '{0}' unknown in usp_update_node_counter",
                                                              mstrNodeName))
              Case 0
                ApplicationLogging.WriteLogEntry(
                  String.Format("Updated transaction counter to database, added '{0}' transactions.", mlngCurrentCounter))
                mlngCurrentCounter = 0
            End Select

          End Using
        End Using


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>Saves the current count value to the counter file.</summary>
    Public Sub SaveCurrentCountToFile()

      Try

        If mblnIsDirty Then
          Using lobjFileStream As New FileStream(FilePath, FileMode.OpenOrCreate)
            Using lobjWriter As New StreamWriter(lobjFileStream)
              lobjWriter.WriteLine(TransactionCounter.EncryptCount(mlngCurrentCount))
              lobjWriter.Close()
            End Using
          End Using
          mblnIsDirty = False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjIdentifierBuilder As New StringBuilder

        lobjIdentifierBuilder.AppendFormat("CurrentCount: {0:n0}", TransactionCounter.CurrentCount)

        If mblnIsDirty Then
          lobjIdentifierBuilder.AppendFormat(" Dirty, <SavedCount: {0:n0}>", TransactionCounter.SavedCount)
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    ''' <summary>Called to raise the PropertyChanged event.</summary>
    Private Sub OnPropertyChanged(ByVal sProp As String)

      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sProp))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Function GetCurrentCount() As Long
      Try
        Dim lintCurrentDBCount As Long = GetCurrentCountFromDB()
        Dim lintCurrentFileCount As Long = 0

        If CounterFileAccessAllowed Then
          lintCurrentFileCount = GetCurrentCountFromFile()
          If lintCurrentFileCount.CompareTo(lintCurrentDBCount) > 0 Then
            Return lintCurrentFileCount
          Else
            Return lintCurrentDBCount
          End If
        Else
          Return lintCurrentDBCount
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetCurrentCountFromDB() As Long
      Try
        If String.IsNullOrEmpty(mstrCatalogConnectionString) Then
          ' We either have not yet initialized the catalog or we have lost it somehow
          ApplicationLogging.WriteLogEntry(
            "Unable to get current count from DB, the connection string has not yet been initialized or has been lost.",
            MethodBase.GetCurrentMethod(), TraceEventType.Warning, 32874)
          Return 0
        End If
        Dim lobjSQLBuilder As New StringBuilder
        lobjSQLBuilder.AppendFormat("SELECT ItemsProcessed FROM tblNodes WHERE NodeAddress = '{0}'", mstrNodeName)
        Dim lobjItemsProcessed As Object = ExecuteSimpleQuery(lobjSQLBuilder.ToString(), mstrCatalogConnectionString)

        If Not IsDBNull(lobjItemsProcessed) Then
          Return CLng(lobjItemsProcessed)
        Else
          Return 0
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Gets the current counter value.</summary>
    Private Shared Function GetCurrentCountFromFile() As Long

      Try

        Dim lstrFirstLine As String = Nothing

        Using lobjFileStream As New FileStream(FilePath, IO.FileMode.Open)
          Using lobjReader As New StreamReader(lobjFileStream)
            lstrFirstLine = lobjReader.ReadLine
            lobjReader.Close()
          End Using
        End Using

        If Not String.IsNullOrEmpty(lstrFirstLine) Then
          Return Instance.DecryptCount(lstrFirstLine)
        End If

      Catch AccessEx As UnauthorizedAccessException
        ApplicationLogging.LogException(AccessEx, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry(
          "Disabling file based counter, access to the counter path is not permitted in this environment.",
          MethodBase.GetCurrentMethod, TraceEventType.Warning, 43876)
        mblnCounterFileAccessAllowed = False
        Return 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return 0
      End Try
    End Function

    ''' <summary>Gets the fully qualified path of the counter file.</summary>
    Private Shared Function GetPrimaryFilePath() As String

      Try

        ' We are using a folder under Microsoft to hide it from prying users.
        ' PTC is for 'Project Transaction Counter'
        Dim lstrCtsProgramDataPath As String = String.Format("{0}\{1}",
                                                             Environment.GetFolderPath(
                                                               SpecialFolder.CommonApplicationData), PARENT_FOLDER_NAME)

        Dim lstrTargetPath As String = Helper.RemoveExtraBackSlashesFromFilePath(lstrCtsProgramDataPath)

        If Directory.Exists(lstrTargetPath) = False Then
          Directory.CreateDirectory(lstrTargetPath)
        End If

        lstrTargetPath = String.Format("{0}\{1}", lstrTargetPath, FileName)

        Return lstrTargetPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Encrypts the count value.</summary>
    ''' <param name="lpCount">The numeric counter value to encrypt.</param>
    ''' <remarks>We encrypt the counter value and save the encrypted string in the counter file.</remarks>
    Private Shared Function EncryptCount(lpCount As Long) As String

      Try

        mobjEncryptedData = mobjSymmetric.Encrypt(New Encryption.Data(lpCount.ToString), mobjKey)
        Return mobjEncryptedData.ToBase64

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Decrypts the count value.</summary>
    ''' <param name="lpEncryptedCount">The encrypted string holding the counter value.</param>
    ''' <remarks>We decrypt the encrypted counter value from the counter file when we open it.</remarks>
    Private Shared Function DecryptCount(lpEncryptedCount As String) As Long

      Try

        mobjEncryptedData.Base64 = lpEncryptedCount
        Using lobjDecryptedData As Encryption.Data = mobjSymmetric.Decrypt(mobjEncryptedData, mobjKey)
          Return CLng(lobjDecryptedData.ToString)
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ExecuteSimpleQuery(ByVal lpSQL As String, ByVal lpConnectionString As String) As Object

      Dim lobjResult As Object

      Try
        Using lobjConnection As New SqlConnection(lpConnectionString)

          Using lobjCommand As New SqlCommand(lpSQL, lobjConnection)
            lobjCommand.CommandTimeout = COMMAND_TIMEOUT

            Helper.HandleConnection(lobjConnection)
            lobjResult = lobjCommand.ExecuteScalar
          End Using
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      End Try

      Return lobjResult
    End Function

    Private Shared Function ExecuteNonQuery(ByVal lpSQL As String, lpConnectionString As String) As Integer

      Dim lintNewRecordAff As Integer = 0

      Try
        Using lobjConnection As New SqlConnection(lpConnectionString)

          Using cmdAdd As New SqlCommand(lpSQL, lobjConnection)
            cmdAdd.CommandTimeout = COMMAND_TIMEOUT

            Helper.HandleConnection(lobjConnection)
            lintNewRecordAff = cmdAdd.ExecuteNonQuery()
          End Using
        End Using
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      End Try

      Return lintNewRecordAff
    End Function

    ''' <summary>Processes the timer event.  Used to save the counter file on a timed basis.</summary>
    Private Sub ProcessTimerEvent(sender As Object,
                                  e As ElapsedEventArgs)

      Try

        If mblnIsDirty Then
          mobjTimer.Enabled = False
          SaveCurrentCount()
          mobjTimer.Enabled = True
        Else
          mlngCurrentCount = GetCurrentCount()
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDisposable Implementation"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)

      Try

        If Not Me.disposedValue Then

          If disposing Then
            ' DISPOSETODO: dispose managed state (managed objects).
            SaveCurrentCount()
          End If

          ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
          ' DISPOSETODO: set large fields to null.
        End If

        Me.disposedValue = True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() _
      Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region
  End Class
End Namespace