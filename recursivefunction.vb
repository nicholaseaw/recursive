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
  
Private Sub RunScript(ByVal C As Line, ByVal r As Double, ByVal n As Double, ByVal A As Double, ByVal L As Double, ByRef RecursiveCurves As Object) 

    'create list of lines
    Dim AllLines As New List (Of Line)

    'call recursive sub
    Call ScaleAndRotate(C, AllLines, A, L)

    'output
    RecursiveCurves = AllLines

End Sub 

Sub ScaleAndRotate(ByRef myLine As Line, ByRef AllLines As List( Of Line), ByVal angle As Double, ByVal MinLength As Double)

    'if polygon is too small, stop looping
    If myLine.Length < MinLength Then Exit Sub

    'make shorter line
    Dim newLine As New Line(myLine.From, myLine.To)

    'rotate
    Dim rotLine As New LineCurve(newLine)

    rotLine.Rotate(angle, Vector3d.ZAxis, Point3d.Origin)

    'calculate scale factor
    Dim scaleLine As New LineCurve(rotLine)
    Dim scaleFactor As Double

    scaleFactor = ( newLine.Length / ( Math.Cos(angle) * ( Math.Tan(angle) + 1) ) ) / newLine.Length

    'apply scale factor
    scaleLine.Scale(scaleFactor)

    'update new line
    newLine = scaleLine.Line

    'update list
    AllLines.Add(newLine)

    'run recursive function

    Call ScaleAndRotate(newLine, AllLines, angle, MinLength)

End Sub
