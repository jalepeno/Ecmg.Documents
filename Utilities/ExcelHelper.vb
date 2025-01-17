'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ExcelHelper.vb
'   Description :  [type_description_here]
'   Created     :  5/28/2013 8:56:21 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports System.IO
Imports OfficeOpenXml
Imports ExcelPackage = OfficeOpenXml.ExcelPackage
Imports ExcelWorksheet = OfficeOpenXml.ExcelWorksheet

#End Region

Namespace Utilities

  Public Class ExcelHelper

    Public Shared Function CreateSpreadSheetFromDataTable(lpSheetName As String, lpDataTable As DataTable) As Object
      Try
        Dim lobjPackage As New ExcelPackage()

        Dim lobjWorkSheet As ExcelWorksheet = lobjPackage.Workbook.Worksheets.Add(lpSheetName)

        lobjWorkSheet.Cells("A1").LoadFromDataTable(lpDataTable, True, Table.TableStyles.Medium9)
        lobjWorkSheet.Cells(lobjWorkSheet.Dimension.Address).AutoFitColumns()

        Return lobjPackage

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateSpreadSheetStreamFromDataTable(lpSheetName As String, lpDataTable As DataTable) As MemoryStream
      Try
        Dim lobjPackage As ExcelPackage = CreateSpreadSheetFromDataTable(lpSheetName, lpDataTable)
        Dim lobjReturnStream As New MemoryStream
        lobjPackage.SaveAs(lobjReturnStream)
        Return lobjReturnStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub CreateSpreadSheetFromDataTable(lpDestinationFilePath As String, lpSheetName As String, lpDataTable As DataTable)
      Try
        Dim lobjSpreadsheetStream As MemoryStream = CreateSpreadSheetStreamFromDataTable(lpSheetName, lpDataTable)
        Helper.WriteStreamToFile(lobjSpreadsheetStream, lpDestinationFilePath, FileMode.Create)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace
