import pythoncom
import win32com.client


def sheetmetal_flat(path, material, relief, thickness, offset, swYearLastDigit = 9):
    sw = win32com.client.Dispatch("SldWorks.Application.{}".format((20+(swYearLastDigit-2))))
    Model = sw.ActiveDoc
    ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    swSelMgr = Model.SelectionManager
    nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)
    partlist = list()
    filename = "files_not_dxf.txt"
    failed_list = open(f"{path}\{filename}", "w+")
    for n in range(1, nbrSelections + 1):
        swComp = swSelMgr.GetSelectedObjectsComponent3(n, 0)
        Component_path = swComp.GetPathName
        if "FILLET_WELD" in str(swComp.Name2):
            continue
        else:
            partlist.append(Component_path)
            print(Component_path)
    for n in range(1, nbrSelections + 1):
            swComp = swSelMgr.GetSelectedObjectsComponent3(n, 0)
            Component_path = swComp.GetPathName
            if Component_path in partlist:
                swComp.GetModelDoc
                number_of_sheet = partlist.count(Component_path)
                partlist = list(filter(lambda a: a != Component_path, partlist))
                sheetmetal_to_dxf(Component_path, path, material,
                                  relief, thickness,
                                  offset, number_of_sheet,
                                  failed_list)


def sheetmetal_to_dxf(Component_path, path, material, relief, thickness,
                      offset, number_of_sheet, failed_list, swYearLastDigit = 9):
    sw = win32com.client.Dispatch("SldWorks.Application.{}".format((20+(swYearLastDigit-2))))

    swDocSpecification = sw.GetOpenDocSpec(Component_path)
    swDocSpecification.FileName
    swDocSpecification.LoadModel = True
    swDocSpecification.UseLightWeightDefault = False
    swDocSpecification.ReadOnly = False
    swDocSpecification.Silent = False

    Model = sw.OpenDoc7(swDocSpecification)
    Model.Visible = True
    sw.ActiveDoc

    if not Model.GetType == 1:
        sw.CloseDoc(Model.GetTitle)

    sw.Callback("fworks@FWPlaybackManager", 0, "SetAutomaticFeatureRecognition 1")
    sw.Callback("fworks@FWPlaybackManager", 0, "RecognizeFeatures")
    sw.Callback("fworks@FWPlaybackManager", 0, "SetAutoDimOptions 0 0 0 0")
    sw.Callback("fworks@FWPlaybackManager", 0, "SetAutoRelOptions 0")

    Model = sw.ActiveDoc
    myModelView = Model.ActiveView
    myModelView.FrameLeft = 0
    myModelView.FrameTop = 22
    Model.ActiveView

    vBodies = Model.GetBodies2(0, True)

    try:
        vBodies[0]
    except:
        sw.CloseDoc(Model.GetTitle)
        return None

    swBody = vBodies[0]
    vFaces = swBody.GetFaces()
    face = vFaces[0]

    for i in range(0, len(vFaces)):
        vFaces[i].SetFaceId(str(i))
        if face.GetArea < vFaces[i].GetArea:
            face = vFaces[i]

    swSelMgr = Model.SelectionManager
    swSelData = swSelMgr.CreateSelectData

    face.Select4(False, swSelData)

    Model.InsertBends2(0, "", 0.5, -1, relief, offset / 1000, True)
    Model.ForceRebuild3(True)

    try:
        if thickness:
            firstfeature = Model.FirstFeature
            nextfeature = firstfeature.GetNextFeature
            while nextfeature.GetTypeName != "SheetMetal":
                firstfeature = nextfeature.GetNextFeature
                nextfeature = firstfeature
            FeatureDefinition = nextfeature.GetDefinition
            thickness = round(1000 * FeatureDefinition.Thickness, 1)
        else:
            thickness = "_"
    except:
        thickness = "_"
        pass

    thickness = str(thickness).replace(".", ",")

    Model = sw.ActiveDoc
    Model.ActiveView

    bit_info = int("000000011111", 2)
    da = [0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0]
    data_align = win32com.client.VARIANT(pythoncom.VT_VARIANT, da)
    data_bit = win32com.client.VARIANT(pythoncom.VT_VARIANT, bit_info)
    varViews = win32com.client.VARIANT(pythoncom.VT_VARIANT, None)
    offset = str(offset).replace(".", ",")
    sPathName = f'{path}\{Model.GetTitle[:-4]}_{material}_thick{thickness}_offset{offset}_quantity{number_of_sheet}.dxf'
    stat = Model.ExportToDWG2(sPathName,
                              Model.GetPathName,
                              1,
                              True,
                              data_align,
                              False,
                              False,
                              data_bit,
                              varViews)
    print(sPathName, stat)
    if not stat:
        failed_list.write("\n"+Model.GetTitle)
    sw.CloseDoc(Model.GetTitle)
