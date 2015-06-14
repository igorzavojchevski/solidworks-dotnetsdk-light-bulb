Imports SldWorks
Imports SwConst

Module Module3

    Sub Assembly(a As SldWorks.SldWorks, pb As ProgressBar)

        Dim app As SldWorks.SldWorks = a
        Dim Part As ModelDoc2
        Dim boolstatus As Boolean
        Dim longstatus As Long
        Dim longwarnings As Long
        Dim myMate As Object
        Dim swAssemblyDoc As AssemblyDoc

        Dim templateFileName As String
        Dim part1 As String
        Dim part2 As String

        'Open Part1
        Part = app.OpenDoc6(CurDir() & "\saved\DolenDel.SLDPRT", 1, 0, "Default", longstatus, longwarnings)
        app.ActivateDoc2("DolenDel.SLDPRT", False, longstatus)

        pb.Increment(10)

        'Open Part2
        Part = app.OpenDoc6(CurDir() & "\saved\GorenDel.SLDPRT", 1, 0, "Default", longstatus, longwarnings)
        app.ActivateDoc2("GorenDel.SLDPRT", False, longstatus)

        pb.Increment(10)

        'Template za Assembly
        templateFileName = CurDir() & "\templates\Assembly.asmdot"

        'New document i setiranje za aktiven
        Part = app.NewDocument(templateFileName, 0, 0, 0)
        app.ActivateDoc2("Assem1", False, longstatus)
        Part = app.ActiveDoc

        pb.Increment(10)

        Dim myModelView As Object
        myModelView = Part.ActiveView
        myModelView.FrameState = swWindowState_e.swWindowMaximized

        'Stavanje na patekite na delovite vo promenlivi
        part1 = CurDir() & "\saved\DolenDel.SLDPRT"
        part2 = CurDir() & "\saved\GorenDel.SLDPRT"

        pb.Increment(10)

        'Dodavanje na delot (part1 - Dolniot del) vo Assembly-to (pritoa mora da se otvori, setira za aktiven i zatvori)
        '=
        boolstatus = Part.AddComponent(part1, 0.0179530997130041, 0.0126836861403064, 0.0262458378498015)
        Part = app.ActiveDoc
        app.ActivateDoc2("DolenDel.SLDPRT", False, longstatus)
        app.CloseDoc("DolenDel.SLDPRT")
        '/=

        pb.Increment(10)

        'Setiranje za aktiven Assem
        app.ActivateDoc2("Assem1", False, longstatus)
        Part = app.ActiveDoc

        'Dodavanje na delot (part2 - Gorniot del) vo Assembly-to (pritoa mora da se otvori, setira za aktiven i zatvori)
        '=
        boolstatus = Part.AddComponent(part2, -0.0330925654852763, 0.0452250542584807, 0.0370982451131567)
        Part = app.ActiveDoc
        app.ActivateDoc2("GorenDel.SLDPRT", False, longstatus)
        app.CloseDoc("GorenDel.SLDPRT")
        '/=

        pb.Increment(10)

        'Setiranje za aktiven Assem
        app.ActivateDoc2("Assem1", False, longstatus)
        Part = app.ActiveDoc

        'Selektiranje na odredeni segmenti, mate i rebuild
        '=
        boolstatus = Part.Extension.SelectByID2("", "FACE", -0.025612064858251, -0.00517340303662195, 0.0427474639297998, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.0193555372106573, 0.00618368614021847, 0.0154700321276664, True, 0, Nothing, 1)
        swAssemblyDoc = Part
        myMate = swAssemblyDoc.AddMate4(4, 1, False, 0.001, 0, 0, 0.001, 0.001, 0.5235987755983, 0.5235987755983, 0.5235987755983, False, False, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()
        '/=

        pb.Increment(10)

        'Selektiranje na soodvetnite segmenti, mate i rebuild
        '=
        boolstatus = Part.Extension.SelectByID2("", "FACE", -0.0272680665337361, 0.0137253807393236, 0.0518824702111829, True, 0, Nothing, 1)
        boolstatus = Part.Extension.SelectByID2("", "FACE", 0.0140884524706166, 0.0254083210713247, 0.0138254786434686, True, 0, Nothing, 1)
        myMate = swAssemblyDoc.AddMate4(1, 1, False, 0.0189755426932127, 0, 0, 0.001, 0.001, 2.07488594482785, 0.5235987755983, 0.5235987755983, False, False, longstatus)
        Part.ClearSelection2(True)
        Part.EditRebuild3()
        '/=

        pb.Increment(10)

        'Dodavanje appearance i rebuild (Ne mozhe da se renderira, bidejki nemam pristap do klasata za PhotoView)
        '=
        Dim strName As String
        Dim swAppearance As RenderMaterial
        Dim swModelDocExt As ModelDocExtension
        Dim nDecalID As Long

        boolstatus = Part.Extension.SelectByID2("GorenDel-1@Assem2", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        swModelDocExt = Part.Extension
        'Appearance on the model
        strName = CurDir() & "\Materials\Glass\gloss\clear glass.p2m"
        swAppearance = swModelDocExt.CreateRenderMaterial(strName)
        boolstatus = swAppearance.AddEntity(Part)
        boolstatus = swModelDocExt.AddRenderMaterial(swAppearance, nDecalID)
        Part.ClearSelection2(True)
        Part.EditRebuild3()
        '/=

        pb.Increment(10)

        'Zacuvuvanje na sklopot (svetilkata)
        longstatus = Part.SaveAs3(CurDir() & "\saved\Assem.SLDASM", 0, 2)

        pb.Increment(10)
    End Sub

End Module

