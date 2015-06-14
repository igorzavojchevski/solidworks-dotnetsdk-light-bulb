Imports SldWorks
Imports SwConst


Module Module1

    Sub GorenDel(a As SldWorks.SldWorks, pb As ProgressBar)

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
        app.ActivateDoc2("GorenDel", False, longstatus)
        Part = app.ActiveDoc

        Dim myModelView As Object
        myModelView = Part.ActiveView
        myModelView.FrameState = swWindowState_e.swWindowMaximized

        pb.Increment(10)

        'Selektiranje na ramnina (Top Plane)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        'Sketch na selektiranata ramnina
        Part.SketchManager.InsertSketch(True)

        'Clear Selection
        Part.ClearSelection2(True)

        'Krug so radius 5mm so centar 7.5mm od centarot
        skSegment = Part.SketchManager.CreateCircleByRadius(0.0075, 0, 0, 0.0005)

        'Clear Selection
        Part.ClearSelection2(True)

        'Exit sketch
        Part.SketchManager.InsertSketch(True)

        'Selektiranje na krugot
        boolstatus = Part.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 4, Nothing, 0)

        'Extrude na krugot (50 mm) 
        myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)

        'Selektiranje na ramnina (Top Plane)
        boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

        'Sketch na selektiranata ramnina
        Part.SketchManager.InsertSketch(True)

        'Clear Selection
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Krug so radius 5mm so centar 7.5mm od centarot (od sprotivnata strana)
        skSegment = Part.SketchManager.CreateCircleByRadius(-0.0075, 0, 0, 0.0005)
        'Extrude na krugot (50 mm) 
        myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0.05, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False


        'boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        'boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        'Sketch na selektiranata ramnina
        Part.SketchManager.InsertSketch(True)
        'Krug
        skSegment = Part.SketchManager.CreateCircle(-0.0#, 0.049097, 0.0#, 0.000006, 0.049049, 0.0#)
        'Extrude
        myFeature = Part.FeatureManager.FeatureExtrusion2(False, False, False, 0, 0, 0.0077, 0.0077, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)
        Part.SelectionManager.EnableContourSelection = False
        'Selektiranje na ramnina (Front Plane)
        boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        'Sketch na selektiranata ramnina
        Part.SketchManager.InsertSketch(True)
        'Krug
        skSegment = Part.SketchManager.CreateCircleByRadius(0, 0.06, 0, 0.03)

        pb.Increment(10)

        Dim vSkLines As Object
        'Corner rectangle
        vSkLines = Part.SketchManager.CreateCornerRectangle(-0.0174728427956549, 0.06, 0, 0.0186001229760197, 0.00994158639065383, 0)
        'Corner rectangle
        vSkLines = Part.SketchManager.CreateCornerRectangle(-0.0131891781102685, 0.0234689485550318, 0, 0.0147673703627793, -0.0112512810002049, 0)
        'Clear selection
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Selektiranje na odredeni segmenti i pravenje trim
        '=
        boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", -0.0120220193660626, 0.06, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", -0.0174728427956549, 0.0520984302435442, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0.0186001229760197, 0.0548604100203884, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0.0130625243530359, 0.0329931405467732, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line5", "SKETCHSEGMENT", -0.0107797531068743, 0.0234689485550318, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line6", "SKETCHSEGMENT", -0.0131891781102685, 0.0179416739109357, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 0.0147673703627793, 0.0197171626579503, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -0.0105261118573008, 0.00994158639065383, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        '/=

        pb.Increment(10)

        'Selektiranje na odredeni tochki od sketch-ot i pravenje sketch fillet
        '=
        boolstatus = Part.Extension.SelectByID2("Point11", "SKETCHPOINT", -0.0174728427956549, 0.0356135331661528, 0, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", 0.0186001229760197, 0.036462042882252, 0, True, 0, Nothing, 1)
        skSegment = Part.SketchManager.CreateFillet(0.02, 1)

        boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -0.0174728427956549, 0.00994158639065383, 0, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("Point4", "SKETCHPOINT", 0.0186001229760197, 0.00994158639065383, 0, True, 0, Nothing, 1)
        skSegment = Part.SketchManager.CreateFillet(0.0025, 1)
        skSegment = Part.SketchManager.CreateFillet(0.0025, 1)

        boolstatus = Part.Extension.SelectByID2("Point13", "SKETCHPOINT", -0.0131891781102685, 0.00994158639065383, 0, False, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("Point14", "SKETCHPOINT", 0.0147673703627793, 0.00994158639065383, 0, True, 0, Nothing, 1)
        skSegment = Part.SketchManager.CreateFillet(0.001, 1)
        skSegment = Part.SketchManager.CreateFillet(0.001, 1)

        boolstatus = Part.Extension.SelectByID2("Point9", "SKETCHPOINT", -0.0131891781102685, -0.0112512810002049, 0, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("Point8", "SKETCHPOINT", 0.0147673703627793, -0.0112512810002049, 0, True, 0, Nothing, 1)
        skSegment = Part.SketchManager.CreateFillet(0.001, 1)
        skSegment = Part.SketchManager.CreateFillet(0.001, 1)
        '/=

        pb.Increment(10)

        'Centralna linija
        skSegment = Part.SketchManager.CreateCenterLine(0.0#, 0.1, 0.0#, 0.0#, -0.02, 0.0#)
        'Clear selection
        Part.ClearSelection2(True)

        'Selektiranje na odredeni segmenti i pravenje trim
        '=
        boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0.0190886110541778, 0.083143572067041, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Arc2", "SKETCHSEGMENT", 0.0214661206502503, 0.038535318942625, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0.0186001229760197, 0.0221778556592323, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Arc4", "SKETCHSEGMENT", 0.0179376693894536, 0.0107464689379456, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line9", "SKETCHSEGMENT", 0.015964781288076, 0.00994158639065383, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Arc6", "SKETCHSEGMENT", 0.0150021337344067, 0.00958533548975758, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 0.0147673703627793, 0.00600395685493027, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Arc8", "SKETCHSEGMENT", 0.014161504183111, -0.0111703340626941, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line7", "SKETCHSEGMENT", 0.00989560658939581, -0.0112512810002049, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        boolstatus = Part.Extension.SelectByID2("Line7", "SKETCHSEGMENT", -0.00412511019445166, -0.0112512810002049, 0, False, 2, Nothing, 0)
        boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0)
        '/=

        pb.Increment(10)

        'Selekcija i brishenje na odreden segment
        boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -0.0145151362264015, 0.0100003123513341, 0, False, 0, Nothing, 0)
        Part.EditDelete()

        'Kreiranje linija (so pochetna tochka na pochetnoto mesto kade bila izbrishana prethodnata, no so pogolema dimenzija)
        skSegment = Part.SketchManager.CreateLine(-0.014973, 0.009942, 0.0#, -0.009973, 0.009942, 0.0#)

        'Clear Selection
        Part.ClearSelection2(True)

        'Selektiranje na krajnata tochka na kreiranata linija i drugata pochetna linija so cel da se napravi constraint - Merge
        boolstatus = Part.Extension.SelectByID2("Point45", "SKETCHPOINT", -0.00997284279565488, 0.00994158639065383, 0, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("Point32", "SKETCHPOINT", -0.0141891781102685, 0.00994158639065383, 0, True, 0, Nothing, 1)
        Part.SketchAddConstraints("sgMERGEPOINTS")
        'Clear Selection
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Selektiranje na odreden segment za da se napravi offset
        boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", -0.023398237343718, 0.0789287422285987, 0, False, 1, Nothing, 0)
        boolstatus = Part.SketchManager.SketchOffset(0.0006, False, True, False, False, True)
        'Clear Selection
        Part.ClearSelection2(True)

        'Kreiranje linija so koja ke se zatvori delot pomegju offsetot (od gornata strana)
        skSegment = Part.SketchManager.CreateLine(0.0#, 0.0906, 0.0#, 0.0#, 0.09, 0.0#)
        'Clear Selection
        Part.ClearSelection2(True)

        'Kreiranje linija so koja ke se zatvori delot pomegju offsetot (od dolnata strana)
        skSegment = Part.SketchManager.CreateLine(-0.007973, -0.011251, 0.0#, -0.007973, -0.011851, 0.0#)
        'Clear Selectio
        Part.ClearSelection2(True)

        pb.Increment(10)

        'Revolve
        myFeature = Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
        Part.SelectionManager.EnableContourSelection = False

        'Zachuvuvanje na part-ot
        longstatus = Part.SaveAs3(CurDir() & "\saved\GorenDel.SLDPRT", 0, 2)

        pb.Increment(10)

    End Sub

End Module