'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ByteStringFactory.vb
'   Description :  [type_description_here]
'   Created     :  2/5/2015 1:30:33 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Files
  Public Class ByteStringFactory

    Friend Shared Function Create(lpString As String) As ByteString
      Try
        Dim lstrByteString As ByteString
        lstrByteString.data = New Byte(((Strings.Len(lpString) - 1) + 1) - 1) {}
        Dim lintItemLength As Integer = (Strings.Len(lpString) - 1)
        Dim lintByteCounter As Integer = 0

        Do While (lintByteCounter <= lintItemLength)
          lstrByteString.data(lintByteCounter) = CByte(Strings.Asc(lpString.Chars(lintByteCounter)))

          If (lstrByteString.data(lintByteCounter) = &H27) Then
            lstrByteString.data(lintByteCounter) = 0
          End If

          lintByteCounter += 1
        Loop

        ByteAlphaToUpper(lstrByteString.data)

        Return lstrByteString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Sub ByteAlphaToUpper(ByRef lpBlock As Byte())
      Dim lintBlockSize As Integer = (lpBlock.Length - 1)
      Dim lintByteCounter As Integer = 0
      Try

        Do While (lintByteCounter <= lintBlockSize)

          Dim lobjByte As Byte = lpBlock(lintByteCounter)

          If ((lobjByte > &H60) And (lobjByte < &H7B)) Then
            lpBlock(lintByteCounter) = CByte((lobjByte - &H20))
          End If

          lintByteCounter += 1
        Loop
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace