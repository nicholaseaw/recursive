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



''' <summary>
''' This class will be instantiated on demand by the Script component.
''' </summary>
Public Class Script_Instance
  Inherits GH_ScriptInstance

  #Region "Utility functions"
  ''' <summary>Print a String to the [Out] Parameter of the Script component.</summary>
  ''' <param name="text">String to print.</param>
  Private Sub Print(ByVal text As String)
    __out.Add(text)
  End Sub
  ''' <summary>Print a formatted String to the [Out] Parameter of the Script component.</summary>
  ''' <param name="format">String format.</param>
  ''' <param name="args">Formatting parameters.</param>
  Private Sub Print(ByVal format As String, ByVal ParamArray args As Object())
    __out.Add(String.Format(format, args))
  End Sub
  ''' <summary>Print useful information about an object instance to the [Out] Parameter of the Script component. </summary>
  ''' <param name="obj">Object instance to parse.</param>
  Private Sub Reflect(ByVal obj As Object)
    __out.Add(GH_ScriptComponentUtilities.ReflectType_VB(obj))
  End Sub
  ''' <summary>Print the signatures of all the overloads of a specific method to the [Out] Parameter of the Script component. </summary>
  ''' <param name="obj">Object instance to parse.</param>
  Private Sub Reflect(ByVal obj As Object, ByVal method_name As String)
    __out.Add(GH_ScriptComponentUtilities.ReflectType_VB(obj, method_name))
  End Sub
#End Region
  
#Region "Members"
  ''' <summary>Gets the current Rhino document.</summary>
  Private RhinoDocument As RhinoDoc
  ''' <summary>Gets the Grasshopper document that owns this script.</summary>
  Private GrasshopperDocument as GH_Document
  ''' <summary>Gets the Grasshopper script component that owns this script.</summary>
  Private Component As IGH_Component
  ''' <summary>
  ''' Gets the current iteration count. The first call to RunScript() is associated with Iteration=0.
  ''' Any subsequent call within the same solution will increment the Iteration count.
  ''' </summary>
  Private Iteration As Integer
#End Region

  ''' <summary>
  ''' This procedure contains the user code. Input parameters are provided as ByVal arguments, 
  ''' Output parameter are ByRef arguments. You don't have to assign output parameters, 
  ''' they will have default values.
  ''' </summary>
  Private Sub RunScript(ByVal C As Line, ByVal A As Double, ByVal S As Double, ByVal L As Double, ByRef Curves As Object)
      'create list of lines
    Dim AllLines As New List ( Of Line )

    'run recursive function
    Call ScaleAndRotate(C, AllLines, A, S, L)

    'output
    Curves = AllLines
  End Sub

  '<Custom additional code> 
    Sub ScaleAndRotate(ByVal myLine As Line, ByRef AllLines As List ( Of Line), ByVal angle As Double, ByVal scale As Double, ByVal MinLength As Double)

    'if square is too small, stop looping
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
    Call ScaleAndRotate(newLine, AllLines, angle, scaleFactor, MinLength)



  End Sub
  '</Custom additional code> 

  Private __err As New List(Of String)
  Private __out As New List(Of String)
  Private doc As RhinoDoc = RhinoDoc.ActiveDoc            'Legacy field.
  Private owner As Grasshopper.Kernel.IGH_ActiveObject    'Legacy field.
  Private runCount As Int32                               'Legacy field.
  
  Public Overrides Sub InvokeRunScript(ByVal owner As IGH_Component, _
                                       ByVal rhinoDocument As Object, _
                                       ByVal iteration As Int32, _
                                       ByVal inputs As List(Of Object), _
                                       ByVal DA As IGH_DataAccess) 
    'Prepare for a new run...
    '1. Reset lists
    Me.__out.Clear()
    Me.__err.Clear()

    'Current field assignments.
    Me.Component = owner
    Me.Iteration = iteration
    Me.GrasshopperDocument = owner.OnPingDocument()
    Me.RhinoDocument = TryCast(rhinoDocument, Rhino.RhinoDoc)

    'Legacy field assignments
    Me.owner = Me.Component
    Me.runCount = Me.Iteration
    Me.doc = Me.RhinoDocument

    '2. Assign input parameters
    Dim C As Line = Nothing
    If (inputs(0) IsNot Nothing) Then
      C = DirectCast(inputs(0), Line)
    End If

    Dim A As Double = Nothing
    If (inputs(1) IsNot Nothing) Then
      A = DirectCast(inputs(1), Double)
    End If

    Dim S As Double = Nothing
    If (inputs(2) IsNot Nothing) Then
      S = DirectCast(inputs(2), Double)
    End If

    Dim L As Double = Nothing
    If (inputs(3) IsNot Nothing) Then
      L = DirectCast(inputs(3), Double)
    End If



    '3. Declare output parameters
  Dim Curves As System.Object = Nothing


    '4. Invoke RunScript
    Call RunScript(C, A, S, L, Curves)

    Try
      '5. Assign output parameters to component...
      If (Curves IsNot Nothing) Then
        If (GH_Format.TreatAsCollection(Curves)) Then
          Dim __enum_Curves As IEnumerable = DirectCast(Curves, IEnumerable)
          DA.SetDataList(1, __enum_Curves)
        Else
          If (TypeOf Curves Is Grasshopper.Kernel.Data.IGH_DataTree) Then
            'merge tree
            DA.SetDataTree(1, DirectCast(Curves, Grasshopper.Kernel.Data.IGH_DataTree))
          Else
            'assign direct
            DA.SetData(1, Curves)
          End If
        End If
      Else
        DA.SetData(1, Nothing)
      End If

    Catch ex As Exception
      __err.Add(String.Format("Script exception: {0}", ex.Message))
    Finally
      'Add errors and messages...
      If (owner.Params.Output.Count > 0) Then
        If (TypeOf owner.Params.Output(0) Is Grasshopper.Kernel.Parameters.Param_String) Then
          Dim __errors_plus_messages As New List(Of String)
          If (Me.__err IsNot Nothing) Then __errors_plus_messages.AddRange(Me.__err)
          If (Me.__out IsNot Nothing) Then __errors_plus_messages.AddRange(Me.__out)
          If (__errors_plus_messages.Count > 0) Then
            DA.SetDataList(0, __errors_plus_messages)
          End If
        End If
      End If
    End Try
  End Sub 
End Class