Option Explicit On

Imports SwConst

Module Module4

    Sub RenderPrev(app As SldWorks.SldWorks)
        Dim Part As SldWorks.ModelDoc2
        Dim swRayTraceRenderer As SldWorks.RayTraceRenderer
        Dim swRayTraceRenderOptions As SldWorks.RayTraceRendererOptions
        Dim errors As Long
        Dim warnings As Long
        Dim status As Boolean

        Part = app.OpenDoc6(CurDir() & "\saved\Assem.SLDASM", 1, 0, "Default", errors, warnings)

        ' Access PhotoView 360
        swRayTraceRenderer = app.GetRayTraceRenderer(1)
        ' Get and set rendering options
        swRayTraceRenderOptions = swRayTraceRenderer.RayTraceRendererOptions
        'Get current rendering values
        Debug.Print("Current rendering values")
        Debug.Print("  ImageHeight          = " & swRayTraceRenderOptions.ImageHeight)
        Debug.Print("  ImageWidth           = " & swRayTraceRenderOptions.ImageWidth)
        Debug.Print("  ImageFormat          = " & swRayTraceRenderOptions.ImageFormat)
        Debug.Print("  PreviewRenderQuality = " & swRayTraceRenderOptions.PreviewRenderQuality)
        Debug.Print("  FinalRenderQuality   = " & swRayTraceRenderOptions.FinalRenderQuality)
        Debug.Print("  BloomEnabled         = " & swRayTraceRenderOptions.BloomEnabled)
        Debug.Print("  BloomThreshold       = " & swRayTraceRenderOptions.BloomThreshold)
        Debug.Print("  BloomRadius          = " & swRayTraceRenderOptions.BloomRadius)
        Debug.Print("  ContourEnabled       = " & swRayTraceRenderOptions.ContourEnabled)
        Debug.Print("  ShadedContour        = " & swRayTraceRenderOptions.ShadedContour)
        Debug.Print("  ContourLineThickness = " & swRayTraceRenderOptions.ContourLineThickness)
        Debug.Print("  ContourLineColor     = " & swRayTraceRenderOptions.ContourLineColor)
        Debug.Print(" ")
        ' Change rendering values
        Debug.Print("New rendering values")
        swRayTraceRenderOptions.ImageHeight = 800
        Debug.Print("  ImageHeight          = " & swRayTraceRenderOptions.ImageHeight)
        swRayTraceRenderOptions.ImageWidth = 800
        Debug.Print("  ImageWidth           = " & swRayTraceRenderOptions.ImageWidth)
        swRayTraceRenderOptions.ImageFormat = 3
        Debug.Print("  ImageFormat          = " & swRayTraceRenderOptions.ImageFormat)
        swRayTraceRenderOptions.PreviewRenderQuality = 1
        Debug.Print("  PreviewRenderQuality = " & swRayTraceRenderOptions.PreviewRenderQuality)
        swRayTraceRenderOptions.FinalRenderQuality = 2
        Debug.Print("  FinalRenderQuality   = " & swRayTraceRenderOptions.FinalRenderQuality)
        swRayTraceRenderOptions.BloomEnabled = False
        Debug.Print("  BloomEnabled         = " & swRayTraceRenderOptions.BloomEnabled)
        swRayTraceRenderOptions.BloomThreshold = 70
        Debug.Print("  BloomThreshold       = " & swRayTraceRenderOptions.BloomThreshold)
        swRayTraceRenderOptions.BloomRadius = 4
        Debug.Print("  BloomRadius          = " & swRayTraceRenderOptions.BloomRadius)
        swRayTraceRenderOptions.ContourEnabled = True
        Debug.Print("  ContourEnabled       = " & swRayTraceRenderOptions.ContourEnabled)
        swRayTraceRenderOptions.ShadedContour = False
        Debug.Print("  ShadedContour        = " & swRayTraceRenderOptions.ShadedContour)
        swRayTraceRenderOptions.ContourLineThickness = 3
        Debug.Print("  ContourLineThickness = " & swRayTraceRenderOptions.ContourLineThickness)
        swRayTraceRenderOptions.ContourLineColor = 255
        Debug.Print("  ContourLineColor     = " & swRayTraceRenderOptions.ContourLineColor)

        'Display the preview window
        status = swRayTraceRenderer.DisplayPreviewWindow
    End Sub
End Module
