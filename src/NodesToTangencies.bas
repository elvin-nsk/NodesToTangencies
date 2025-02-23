Attribute VB_Name = "NodesToTangencies"
'===============================================================================
'   Макрос          : NodesToTangencies
'   Версия          : 2025.02.23
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "NodesToTangencies"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2025.02.23"

'===============================================================================
' # Globals

Private Const SIZE_TO_TOLERANCE_MULT As Double = 0.001

'===============================================================================
' # Entry points

Sub Start()
    MainVariants True
End Sub

Sub StartKeepNodes()
    MainVariants False
End Sub

'===============================================================================
' # Main

Private Sub MainVariants(ByVal DeleteSourceNodes As Boolean)
    
    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Shapes As ShapeRange
    If Not InputData.ExpectShapes.Ok(Shapes) Then GoTo Finally
    
    Dim Source As ShapeRange: Set Source = ActiveSelectionRange
    
    BoostStart APP_DISPLAYNAME
    
    Dim Shape As Shape
    For Each Shape In Shapes
        If HasCurve(Shape) Then
            ProcessCurve Shape.Curve, DeleteSourceNodes
        End If
    Next Shape
    
    If IsSome(Source) Then Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally
    
End Sub

'===============================================================================
' # Helpers

Private Sub ProcessCurve( _
                ByVal Curve As Curve, _
                ByVal DeleteSourceNodes As Boolean _
            )
    Dim Points As New Collection, PointsToKeep As New Collection
    If DeleteSourceNodes Then Curve.Nodes.All.Smoothen 0
    FindTangents Curve, 0, Points, PointsToKeep
    FindTangents Curve, 90, Points, PointsToKeep
    #If DebugMode = 1 Then
    MakeMarks(Points, Blue).OrderToBack
    MakeMarks(PointsToKeep, Green).OrderToBack
    #End If
    If Points.Count = 0 Then Exit Sub
    
    AddNodes Curve, Points
    
    If Not DeleteSourceNodes Then Exit Sub
    
    AppendCollection PointsToKeep, Points
    Dim NodesToDelete As NodeRange: Set NodesToDelete = _
        FindNodesToDelete(Curve, PointsToKeep)
    NodesToDelete.Delete
End Sub

Private Sub FindTangents( _
                ByVal Curve As Curve, _
                ByVal Angle As Double, _
                ByVal PointsPool As Collection, _
                ByVal PointsToKeepPool As Collection _
            )
    Dim Segment As Segment
    Dim Offset1 As Double, Offset2 As Double, n As Long
    For Each Segment In Curve.Segments
        n = Segment.GetPeaks(Angle, Offset1, Offset2, cdrParamSegmentOffset)
        If n > 1 Then SortPoint Segment, Offset2, PointsPool, PointsToKeepPool
        If n > 0 Then SortPoint Segment, Offset1, PointsPool, PointsToKeepPool
    Next Segment
End Sub

Private Sub SortPoint( _
                ByVal Segment As Segment, _
                ByVal ParamOffset As Double, _
                ByVal PointsPool As Collection, _
                ByVal PointsToKeepPool As Collection _
            )
    Dim Point As Point: Set Point = OffsetToPoint(Segment, ParamOffset)
    If ParamOffset = 0 Or ParamOffset = 1 Then
        PointsToKeepPool.Add Point
        Exit Sub
    End If
    If IsNodeMatchPoint(Segment.StartNode, Point) _
    Or IsNodeMatchPoint(Segment.EndNode, Point) Then
        PointsToKeepPool.Add Point
    Else
        PointsPool.Add Point
    End If
End Sub

Private Property Get OffsetToPoint( _
                         ByVal Segment As Segment, _
                         ByVal ParamOffset As Double _
                     ) As Point
    Dim x As Double, y As Double
    Segment.GetPointPositionAt x, y, ParamOffset, cdrParamSegmentOffset
    Set OffsetToPoint = Point.New_(x, y)
End Property

Private Function AddNodes( _
                     ByVal Curve As Curve, _
                     ByVal Points As Collection _
                 ) As NodeRange
    Set AddNodes = CreateNodeRange
    Dim Offset As Double, Segment As Segment
    Dim Point As Point
    For Each Point In Points
        Set Segment = Curve.FindClosestSegment(Point.x, Point.y, Offset)
        AddNodes.Add Segment.AddNodeAt(Offset, cdrParamSegmentOffset)
    Next
End Function

Private Property Get FindNodesToDelete( _
                         ByVal Curve As Curve, _
                         ByVal PointsToKeep As Collection _
                     ) As NodeRange
    Set FindNodesToDelete = CreateNodeRange
    Dim Node As Node
    For Each Node In Curve.Nodes
        If Not Node.IsEnding Then
            If Not Node.Type = cdrCuspNode _
           And Not Node.Segment.Type = cdrLineSegment Then
                If Not IsNodeMatchPoints(Node, PointsToKeep) Then
                    FindNodesToDelete.Add Node
                End If
            End If
        End If
    Next Node
End Property

Private Property Get IsNodeMatchPoints( _
                         ByVal Node As Node, _
                         ByVal Points As Collection _
                     ) As Boolean
    Dim Point As Point
    For Each Point In Points
        If IsNodeMatchPoint(Node, Point) Then
            IsNodeMatchPoints = True
            Exit Property
        End If
    Next Point
End Property

Private Property Get IsNodeMatchPoint( _
                         ByVal Node As Node, _
                         ByVal Point As Point _
                     ) As Boolean
    Dim t As Double: t = Tolerance(Node)
    If IsApproximate(Node.PositionX, Point.x, t) _
   And IsApproximate(Node.PositionY, Point.y, t) Then
        IsNodeMatchPoint = True
    End If
End Property

Private Property Get Tolerance(ByVal Node As Node) As Double
    Dim Segment As Segment: Set Segment = Node.Segment
    If IsNone(Segment) Then Exit Property
    Tolerance = _
        (Node.Segment.BoundingBox.Width + Node.Segment.BoundingBox.Height / 2) _
      * SIZE_TO_TOLERANCE_MULT
End Property

Private Function MakeMarks(ByVal Points As Collection, ByVal Color As Color) As ShapeRange
    Set MakeMarks = CreateShapeRange
    Dim Point As Point
    For Each Point In Points
        MakeMarks.Add MakeCircle(Point.x, Point.y, 1, Color)
    Next Point
End Function

'===============================================================================
' # Tests

Private Sub TestSmooth()
    ActiveSelectionRange.FirstShape.Curve.Nodes.All.Smoothen 0
End Sub
