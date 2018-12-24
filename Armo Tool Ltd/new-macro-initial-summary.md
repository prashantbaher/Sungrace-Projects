# Let's Make summary for below code

I will explain 1 line at a time

```vb
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swDoc As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc

Dim swBody As SldWorks.Body2
Dim swFace As SldWorks.Face2

Dim swLoop As SldWorks.Loop2
Dim swCurve As SldWorks.Curve

Dim swFace1, swFace2, swTheFace As SldWorks.Surface

Dim swEntity As SldWorks.Entity
Dim swSafeEntity As SldWorks.Entity

' Collection of edges
Dim EdgeCollection As Collection
' Collection of curves
Dim CurveCollection As Collection
' Collection of surfaces
Dim SurfaceCollection As Collection

Dim i, j, z As Integer

Public Edge, EdgeParameter1, EdgeParameter2, Faces As Variant


Sub main()
    
    Set EdgeCollection = New Collection
    Set CurveCollection = New Collection
    Set SurfaceCollection = New Collection
    
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    Set swPart = swDoc
        
    Set swBody = swPart.Body
    Set swFace = swBody.GetFirstFace
    
    Do While Not swFace Is Nothing
        
        Set swLoop = swFace.GetFirstLoop
        While Not swLoop Is Nothing
            
            If swLoop.IsOuter = False Then
                
                Edge = swLoop.GetEdges
                For i = 0 To UBound(Edge)
                    Set swCurve = Edge(i).GetCurve
                    
                    If swCurve.Identity = CIRCLE_TYPE Then
                    
                        EdgeCollection.Add Edge(i)
                        CurveCollection.Add swCurve
                        Edge(i).Select True
                    
'                        If Trim(Str(Round(swCurve.CircleParams(6) * 1000 * 2, 10))) = Trim(dd) Then
'                            EdgeCollection.Add Edge(i)
'                            CurveCollection.Add swCurve
'                            Edge(i).Select True
'                        End If
                        
                    End If
                    
                Next i
                
            End If
            
            Set swLoop = swLoop.GetNext
        Wend
        
        Set swFace = swFace.GetNextFace
        
    Loop
    
    z = 0
    i = 1
    
    Do While i <= CurveCollection.Count
    
        EdgeParameter1 = CurveCollection(i).CircleParams
        j = i + 1
        
        Do While j <= CurveCollection.Count
        
            EdgeParameter2 = CurveCollection(j).CircleParams
            
            If (Round(EdgeParameter1(0), 10) = Round(EdgeParameter2(0), 10)) And (Round(EdgeParameter1(1), 10) = Round(EdgeParameter2(1), 10)) Then
                CurveCollection.Remove (j)
                EdgeCollection.Remove (j)
                j = i
            End If
            j = j + 1
        
        Loop
        i = i + 1
        
    Loop
    
    For i = 1 To CurveCollection.Count
        If i = 1 Then
            EdgeCollection(i).Select False
        Else
            EdgeCollection(i).Select True
        End If
    Next
    
    For i = 1 To EdgeCollection.Count
        
        Faces = EdgeCollection.Item(i).GetTwoAdjacentFaces2
        Set swFace1 = Faces(0).GetSurface
        Set swFace2 = Faces(1).GetSurface
        
        If swFace1.IsCylinder Then
            Set swTheFace = swFace1
            Set swEntity = Faces(0)
        Else
            Set swTheFace = swFace2
            Set swEntity = Faces(1)
        End If
        
        If Not swTheFace Is Nothing Then
            
            Set swSafeEntity = swEntity.GetSafeEntity
            swSafeEntity.Select True
            z = z + 1
            SurfaceCollection.Add swSafeEntity
            
        End If
        
    Next
    
    MsgBox (z & " holes are found ")

End Sub


```

---

```vb
Option Explicit
```

This line is for making sure that we did not use any un-defined variable.

Which may cause an error in future.

---

```vb
Dim swApp As SldWorks.SldWorks
```

This line is define our variable `swApp` as *Solidworks* application.

---

```vb
Dim swDoc As SldWorks.ModelDoc2
```

This line is define our variable `swDoc` as *ModelDoc* for our document.

---

```vb
Dim swPart As SldWorks.PartDoc
```

This line is define our variable `swPart` as *Part* document.

---

```vb
Dim swBody As SldWorks.Body2
```

This line is define our variable `swBody` as *Solidworks Body*.

---

```vb
Dim swFace As SldWorks.Face2
```

This line is define our variable `swFace` as *Solidworks Face*.

---

```vb
Dim swLoop As SldWorks.Loop2
```

This line is define our variable `swLoop` as *Solidworks Loop*.

---

```vb
Dim swCurve As SldWorks.Curve
```

This line is define our variable `swCurve` as *Solidworks Curve*.

---

```vb
Dim swFace1, swFace2, swTheFace As SldWorks.Surface
```

This line is define our variables named `swFace1, swFace2, swTheFace` as *Solidworks Surface*.

---

```vb
Dim swEntity As SldWorks.Entity
```

This line is define our variable named `swEntity` as *Solidworks Entity*.

---

```vb
Dim swSafeEntity As SldWorks.Entity
```

This line is define our variable named `swSafeEntity` as *Solidworks Entity* also but the purpose is to collect all **safe entities**.

---

```vb
' Collection of Edges
Dim EdgeCollection As Collection
```

This line is define our variable named `EdgeCollection` as *Collection* of **edges** into our program.

---

```vb
' Collection of curves
Dim CurveCollection As Collection
```

This line is define our variable named `CurveCollection` as *Collection* of **curves** into our program.

---

```vb
' Collection of surfaces
Dim SurfaceCollection As Collection
```

This line is define our variable named `SurfaceCollection` as *Collection* of **Surfaces** into our program.

---

```vb
Dim i, j, z As Integer
```

This line is define *Integer* variable to use in our program.

---

```vb
Public Edge, EdgeParameter1, EdgeParameter2, Faces As Variant
```

This line is define our variable named `Edge, EdgeParameter1, EdgeParameter2, Faces` as *Variant* to use in our program.

---

```vb
Sub main()

End Sub
```

This is a ***function*** which is *starting* function as well.

It is important to note that all the `code` inside the *main function* is executed 1st.

In this code our main function is named as `main()`, which is a common way to define main function.

---

```vb
Set EdgeCollection = New Collection
Set CurveCollection = New Collection
Set SurfaceCollection = New Collection
```

This line is set our variables named `EdgeCollection, CurveCollection, SurfaceCollection` as new *Collection* to use in our program.

---

```vb
Set swApp = Application.SldWorks
```

This line is set our variable named `swApp` as new *Solidworks* application.

---

```vb
Set swDoc = swApp.ActiveDoc
```

This line is set our variable named `swDoc` as new *Solidworks Active* document.

---

```vb
Set swPart = swDoc
```

This line is set our variable named `swPart` as `swDoc` *Solidworks Active* document.

---

```vb
Set swBody = swPart.Body
```

This line is set our variable named `swBody` as **Body** in our part file.

---

```vb
Set swFace = swBody.GetFirstFace
```

This line is use a *method* name `GetFirstFace` of `swBody` object.

This *method* gets the **First face** of body and sets this face to our variable i.e. `swFace`.

> Since `swFace` is defined as `SldWorks.Face2` hence after executing this line `swFace` can use *methods and properties* of `SldWorks.Face2` object.

For more information use ***Solidworks API help*** our your Solidworks software. It will give you more indepth information.

---