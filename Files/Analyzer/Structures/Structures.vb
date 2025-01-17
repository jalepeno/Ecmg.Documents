'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  Structures.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 3:47:57 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Runtime.InteropServices
Imports Newtonsoft.Json

#End Region

Namespace Files

  <Serializable, StructLayout(LayoutKind.Sequential)> Public Structure ByteString
    Public data As Byte()
  End Structure

  '<StructLayout(LayoutKind.Sequential)> Public Structure FileTestResult
  '  Public FileType As String
  '  Public FileExt As String
  '  Public Points As Integer
  '  Public Perc As Single
  '  Public ExtraInfo As ExtraInfo
  'End Structure

  <Serializable, StructLayout(LayoutKind.Sequential)> Public Structure ExtraInfo
    <JsonIgnore()>
    Public FileType As String
    <JsonIgnore()>
    Public FileExt As String
    <JsonIgnore()>
    Public FilePts As String
    Public FilesScanned As Integer
    <JsonIgnore()>
    Public AuthorName As String
    <JsonIgnore()>
    Public AuthorEMail As String
    <JsonIgnore()>
    Public AuthorHome As String
    <JsonIgnore()>
    Public DefFile As String
    Public Remark As String
    Public RelURL As String
  End Structure

End Namespace