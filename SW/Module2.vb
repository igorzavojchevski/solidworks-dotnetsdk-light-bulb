Imports SldWorks
Imports SwConst


Module Module2

    Sub DolenDel(a As SldWorks.SldWorks, pb As ProgressBar)

        Dim app As SldWorks.SldWorks = a
        Dim Part As ModelDoc2
        Dim boolstatus As Boolean
        Dim longstatus As Long
        Dim skSegment As SketchSegment
        Dim myFeature As Feature

        Dim templateFileName As String
        'Selektiranje na template za part
        templateFileName = CurDir() & "\templates\Part.prtdot"
        'Kreiranje na nov dokument
        Part = app.NewDocument(templateFileName, 0, 0, 0)
        'Setiranje na dokumentot kako momentalno aktiven
        app.ActivateDoc2("DolenDel", False, longstatus)
        Part = app.ActiveDoc

        Dim myModelView As Object
        myModelView = Part.ActiveView
        myModelView.FrameState = swWindowState_e.swWindowMaximized

        pb.Increment(10)

        'Selektiranje na ramnina
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        'Sketch na selektiranata ramnina
        Part.SketchManager.InsertSketch(True)

        'Kreiranje krug so radius 13.5mm
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.0135)

        'Extrude na krugot (so depth 22) (za da se dobie cilindar koj ke pretstavuva grlo na svetilkata)
        myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.022, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)

        pb.Increment(10)
        'Selektiranje na ramnina (Right Plane)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        'Sketch na selektiranata (desna) ramnina
        Part.SketchManager.InsertSketch(True)

        'Krug so radius 15mm so centar vo rabot na prethodniot
        skSegment = Part.SketchManager.CreateCircleByRadius(-0.0135, 0, 0, 0.0015)

        'Exit sketch
        Part.SketchManager.InsertSketch(False)

        'Clear Selection
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Selektiranje na ramnina (Top Plane)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        'Sketch na selektiranata (top) ramnina
        Part.SketchManager.InsertSketch(True)

        'Krug so radius 13.5mm
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.0135)

        'Spirala (pitch 4mm, revolutions 4)
        Part.InsertHelix(False, True, False, True, 0, 0.02, 0.004, 4, 0, 6.2831853071796)

        'Clear Selection
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Selektiranje na krugot i spiralata
        boolstatus = Part.Extension.SelectByID2("Sketch2", "SKETCH", -0.01234600704513, 0.000958279844362024, 0, False, 1, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Helix/Spiral1", "REFERENCECURVES", 0.0816497384695936, 0.0547983122366702, 0.112805214880893, True, 4, Nothing, 0)
        'Swept cut - krugot po spiralata
        myFeature = Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False)


        'Selektiranje na gornata osnova
        boolstatus = Part.Extension.SelectByID2("", "FACE", -0.00455727409968176, 0.019999999999925, 0.00108491963959523, False, 1, Nothing, 0)
        'Shell - se pravi otvor (dupka) koja ke pretstavuva vlez za svetilkata (drugiot del)
        Part.InsertFeatureShell(0.0005, False)

        pb.Increment(10)

        'Selektiranje na Face1 i Face2 so cel podocna da se napravi Fillet spored istite
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00315076975266493, 0.0192681794011378, 0.0131271721996882, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00413598729613796, 0.0156911570798854, 0.0112672312876043, True, 4, Nothing, 0)

        Dim radiiArray0 As Object
        Dim radiis0 As Double
        Dim conicRhosArray0 As Object
        Dim coniRhos0 As Double
        Dim setBackArray0 As Object
        Dim setBacks0 As Double
        Dim pointArray0 As Object
        Dim points0 As Double
        Dim pointRhoArray0 As Object
        Dim pointsRhos0 As Double

        radiiArray0 = radiis0
        conicRhosArray0 = coniRhos0
        setBackArray0 = setBacks0
        pointArray0 = points0
        pointRhoArray0 = pointsRhos0

        'Fillet - radius 0.30mm
        myFeature = Part.FeatureManager.FeatureFillet2(195, 0.0003, 0, 2, 0, 0, (radiiArray0), (conicRhosArray0), (setBackArray0), (pointArray0), (pointRhoArray0))

        'Selektiranje obratno, so cel da se napravi fillet i od drugata strana (sega se selektiraat obratno)
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00484876254682831, 0.0156110562669198, 0.0109823400942446, False, 2, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00513766879356581, 0.0194062639933748, 0.0124841643439026, True, 4, Nothing, 0)

        pb.Increment(10)

        Dim radiiArray1 As Object
        Dim radiis1 As Double
        Dim conicRhosArray1 As Object
        Dim coniRhos1 As Double
        Dim setBackArray1 As Object
        Dim setBacks1 As Double
        Dim pointArray1 As Object
        Dim points1 As Double
        Dim pointRhoArray1 As Object
        Dim pointsRhos1 As Double

        radiiArray1 = radiis1
        conicRhosArray1 = coniRhos1
        setBackArray1 = setBacks1
        pointArray1 = points1
        pointRhoArray1 = pointsRhos1

        'Fillet - radius 0.30mm
        myFeature = Part.FeatureManager.FeatureFillet2(195, 0.0003, 0, 2, 0, 0, (radiiArray1), (conicRhosArray1), (setBackArray1), (pointArray1), (pointRhoArray1))

        pb.Increment(10)

        'Se selektira dnoto na modelot
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00380106406096914, 0, -0.0023732776815012, False, 0, Nothing, 0)
        'Sketch na selektiranoto
        Part.SketchManager.InsertSketch(True)
        'Krug so radius 10mm
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.01)
        'ExtrudeCut (up to next) - po krugot, za da se napravi dupka, otvor na dnoto
        myFeature = Part.FeatureManager.FeatureCut3(True, False, False, 2, 0, 0.02, 0.02, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, False, True, True, True, True, False, 0, 0, False)

        pb.Increment(10)

        'Selekcija na dolniot del (okolu otvorot)
        boolstatus = Part.Extension.SelectByID2("", "FACE", -0.00321673074499727, 0, 0.0106268509407023, False, 0, Nothing, 0)
        'Sketch na selektiranoto
        Part.SketchManager.InsertSketch(True)
        'Krug so radius 12.5mm
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.0125)
        'ExtrudeBoss (direction - blind, depth - 2,5mm, draft on - 45 deg)
        myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.0025, 0.02, True, False, False, False, 0.78539816339745, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)
        'Part.SelectionManager.EnableContourSelection = False

        pb.Increment(10)

        'Selektiranje na dnoto od ekstrudiranoto
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00210640978078525, -0.00249999999999773, 0.00402837466395105, False, 0, Nothing, 0)
        'Sketch na selektiranoto
        Part.SketchManager.InsertSketch(True)
        'Krug so radius 10mm
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.01)
        'ExtrudeBoss (blind, depth - 4mm, draft - 45deg)
        myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.004, 0.0025, True, False, False, False, 0.78539816339745, 0.78539816339745, False, False, False, False, True, True, True, 0, 0, False)
        'Part.SelectionManager.EnableContourSelection = False
        'Selektiranje na dnoto od ekstrudiranoto
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.00072017526886272, -0.00650000000001683, 0.00375112776156657, False, 1, Nothing, 0)
        'Dome
        Part.InsertDome(0.0015, False, False)

        'Zachuvuvanje na part-ot
        longstatus = Part.SaveAs3(CurDir() & "\saved\DolenDel.SLDPRT", 0, 2)

        pb.Increment(10)
    End Sub


End Module
