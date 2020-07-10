Option Strict Off
Option Explicit On

Imports System
Imports System.Collections
Imports System.Collections.Generic

Imports Rhino
Imports Rhino.Geometry

Imports Grasshopper
Imports Grasshopper.Kernel
Imports Grasshopper.Kernel.Data
Imports Grasshopper.Kernel.Types

Public Class Script_Instance
  Inherits GH_ScriptInstance

Private Sub RunScript(ByVal r As Double, ByVal n As Double, ByVal A As Double, ByVal L As Double, ByRef C As Object) 
    Dim crv As New Polyline
    Dim origin As New Point3d(0, 0, 0)
    Dim circle As New Circle(origin, r)

    'create n-sided polygon
    crv = Polyline.CreateCircumscribedPolygon(circle, n)

    'get segments of curve
    Dim segments As Line() = Nothing

    segments = crv.GetSegments

    'create list of segments
    Dim lstsegments As New List (Of Line)
    Dim i As Integer

    For i = 0 To n - 1
      lstsegments.Add(segments(i))
    Next i

    'output
    C = lstsegments

  End Sub 

End Class
