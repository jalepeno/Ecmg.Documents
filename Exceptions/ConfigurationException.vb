'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Exceptions

  ''' <summary> 
  ''' This is an exception specifically classed for the SysConfiguration class. 
  ''' See http://msdn2.microsoft.com/en-us/library/dd14ef5c-80e6-41a5-834e-eba8e2eae75e(vs.80).aspx for details. 
  ''' </summary> 
  <System.Serializable()>
  Public Class ConfigurationException
    Inherits CtsException
    ' 
    ' For guidelines regarding the creation of new exception types, see 
    ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp 
    ' and 
    ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp 
    ' 

    Public Sub New()
    End Sub
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub
    Public Sub New(ByVal message As String, ByVal inner As Exception)
      MyBase.New(message, inner)
    End Sub
    Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
      MyBase.New(info, context)
    End Sub
  End Class
End Namespace