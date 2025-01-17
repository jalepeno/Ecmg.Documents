'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Annotations.Redaction

  ''' <summary>
  ''' Marker class and root of any future redaction-type classes.
  ''' </summary>
  <Serializable()>
  Public Class RedactionBase
    Inherits Annotation

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="RedactionBase" /> class.
    ''' </summary>
    Public Sub New()
      Me.IsRedaction = True
    End Sub

#End Region

  End Class

End Namespace