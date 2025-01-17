'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Arguments

  Public Class SimpleSearchArgs

    Private mstrSearchTerm As String = String.Empty
    Private mstrStartFolder As String = String.Empty
    Private mstrQueryTarget As String = String.Empty
    Private mlngMaxResultCount As Long = -1


    Public Property SearchTerm() As String
      Get
        Return mstrSearchTerm
      End Get
      Set(ByVal value As String)
        mstrSearchTerm = value
      End Set
    End Property

    Public Property StartFolder() As String
      Get
        Return mstrStartFolder
      End Get
      Set(ByVal value As String)
        mstrStartFolder = value
      End Set
    End Property

    Public Property QueryTarget() As String
      Get
        Return mstrQueryTarget
      End Get
      Set(ByVal value As String)
        mstrQueryTarget = value
      End Set
    End Property

    Public Property MaxResultCount() As Long
      Get
        Return mlngMaxResultCount
      End Get
      Set(ByVal value As Long)
        mlngMaxResultCount = value
      End Set
    End Property

  End Class

End Namespace