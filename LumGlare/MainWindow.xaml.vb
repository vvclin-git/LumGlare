Imports LightTools
Imports LTCOM64
Imports LTJumpStartNET
Imports LTLOCATORLib
Imports System.ComponentModel
Imports System.Math



Class MainWindow
    Dim ltLoc As New Locator
    Dim lt As LTAPI
    Dim js As New LTCOM64.JSNET
    Dim pID As Integer
    Dim stat As LTReturnCodeEnum
    Dim meshKeyDict As New Dictionary(Of String, String) 'a dictionary for mapping mesh names and keys
    Dim receiverKeyDict As New Dictionary(Of String, String) 'another dictionary for mapping mesh names and its receivers
    Dim posIdArray(,) As Double
    Dim strArray(,) As Double
    'variables for pinhole cam creation
    Dim newDummyPlaneKey As String
    Dim newDummyPlaneName As String
    Dim newCamKey As String
    Dim newCamName As String
    Dim newGroupKey, newGroupName As String
    Dim bw As New BackgroundWorker() 'start a new thread for background process
    Dim stopWatch As New Stopwatch()


    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        bw.WorkerReportsProgress = True
        pID = 7540
        lt = ltLoc.GetLTAPI(pID)
        Call refreshMeshList()

    End Sub
    Private Sub refreshMeshList()
        'retrieve a list of luminance/illuminance meshes in LT
        Dim receiverListKey, receiverKey As String
        Dim meshListKey, meshKey As String
        Dim meshName As String
        Dim statList1 As LTReturnCodeEnum
        Dim statList2 As LTReturnCodeEnum
        'get receiver list
        receiverListKey = lt.DbList("LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager]", "RECEIVER")
        'loop through the receivers
        Do
            receiverKey = lt.ListNext(receiverListKey, statList1)
            'for each receiver, return a list containing all meshes
            meshListKey = lt.DbList(receiverKey, "MESH")
            'use return code for loop ternimation
            If statList1 = 53 Then
                Exit Do
            End If
            'loop through the meshes
            Do
                meshKey = lt.ListNext(meshListKey, statList2)
                If statList2 = 53 Then
                    Exit Do
                End If
                'exclude intensity mesh
                If lt.DbType(meshKey, "INTENSITY_MESH") = 0 Then
                    If lt.DbGet(meshKey, "Name") Like "PinHoleLumMap*" Then
                        meshName = lt.DbGet(receiverKey, "Name") + " | " + lt.DbGet(meshKey, "Name")
                        'put mesh name-key mapping into dictionary
                        meshKeyDict.Add(meshName, meshKey)
                        'put mesh name-receiver key mapping into dictionary
                        receiverKeyDict.Add(meshName, receiverKey)
                        'add meshName to the combobox
                        Me.meshList.Items.Add(meshName)
                    End If
                End If
            Loop
        Loop
    End Sub

    Private Sub refreshMeshListBtn_Click(sender As Object, e As RoutedEventArgs) Handles refreshMeshListBtn.Click
        Call refreshMeshList()
    End Sub


    Private Sub calcUGR_Click(sender As Object, e As RoutedEventArgs) Handles calcUGR.Click
        Dim receiverKey, meshKey As String
        Dim numSample As Long
        Dim lumArray(,) As Double
        Dim ugrArray(,) As Double
        Dim UGR, ugrTmp As Double
        Dim backLum As Double
        Dim numX, numY As Integer
        Dim ltStat As LTReturnCodeEnum
        receiverKey = receiverKeyDict(meshList.Text)
        meshKey = meshKeyDict(meshList.Text)
        backLum = Double.Parse(Me.backLumIn.Text)
        numSample = lt.DbGet(meshKey, "Num_Samples")
        numX = lt.DbGet(meshKey, "X_Dimension")
        numY = lt.DbGet(meshKey, "Y_Dimension")
        If numSample = 0 Then
        Else
            ReDim lumArray(numX - 1, numY - 1)
            ReDim ugrArray(numY - 1, numX - 1)
            ltStat = lt.GetMeshData(meshKey, lumArray, "CellValue")
            ugrTmp = 0
            For i = 0 To numY - 1
                For j = 0 To numX - 1
                    ugrTmp += (lumArray(j, i)) ^ 2 * strArray(i, j) / (posIdArray(i, j)) ^ 2
                    ugrArray(i, j) = (lumArray(j, numY - 1 - i)) ^ 2 * strArray(i, j) / (posIdArray(i, j)) ^ 2
                Next
            Next
            'Debug.Print(ugrTmp)
            UGR = 8 * Log10(0.25 / backLum * ugrTmp)
            Me.ugrOut.Text = String.Format("{0:0.0}", UGR)
            Call TwoDArrayToCSV(lumArray)
            Call TwoDArrayToCSV(ugrArray)
        End If


    End Sub
    Private Sub bw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        'stopWatch.Stop()
        'Dim ts As TimeSpan = stopWatch.Elapsed
        ''Format And display the TimeSpan value.
        'Dim elapsedTime As String = String.Format("{0:0.0000}", ts.TotalSeconds)
        'Me.label1.Content = "Ready (Last Running Time: " + elapsedTime + " secs)"
    End Sub

    Private Function posInd(gridX As Double, gridY As Double, workDist As Double) As Double
        Dim theta, tau As Double
        theta = gridTheta(gridX, gridY, workDist)
        tau = gridTau(gridX, gridY)
        If gridY > 0 Then
            Return guth(tau, theta)
        Else
            Return iwata(theta)
        End If
    End Function
    Private Function guth(tau As Double, theta As Double) As Double
        Return Exp((35.2 - 0.31889 * tau - 1.22 * Exp(-2 * tau / 9)) * 10 ^ -3 * theta + (21 + 0.26667 * tau - 0.002963 * tau ^ 2) * 10 ^ -5 * theta ^ 2)
    End Function
    Private Function iwata(theta As Double) As Double
        Dim mult As Double
        If theta < 30.96375653 Then
            mult = 0.8
        Else
            mult = 1.2
        End If
        Return 1 + mult * Tan(RAD(theta))
    End Function
    Private Function gridStr(theta As Double, gridArea As Double, s As Double, workDist As Double) As Double

        Return gridArea * workDist / s ^ 3
    End Function
    Private Function gridTheta(gridX As Double, gridY As Double, workDist As Double) As Double
        Return DEG(Atan(Sqrt(gridX ^ 2 + gridY ^ 2) / workDist))
    End Function
    Private Function gridTau(gridX As Double, gridY As Double) As Double
        Return Abs(DEG(Atan(gridX / gridY)))
    End Function
    Private Sub initMesh()
        'initialize arrays of position index and solid angle
        Dim receiverKey, meshKey As String
        Dim gridX, gridY As Double
        Dim gridX0, gridY0 As Double
        Dim gridArea As Double
        Dim workDist As Double
        Dim gridWidth, gridHeight As Double
        Dim s As Double
        Dim numX, numY As Integer
        Dim minStr, maxStr As Double
        receiverKey = receiverKeyDict(Me.meshList.Text)
        meshKey = meshKeyDict(Me.meshList.Text)
        workDist = lt.DbGet(lt.DbKeyStr(receiverKey) + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "Distance")
        'get grid number and dimension from mesh
        numX = lt.DbGet(meshKey, "X_Dimension")
        numY = lt.DbGet(meshKey, "Y_Dimension")
        gridWidth = (lt.DbGet(meshKey, "Max_X_Bound") - lt.DbGet(meshKey, "Min_X_Bound")) / numX
        gridHeight = (lt.DbGet(meshKey, "Max_Y_Bound") - lt.DbGet(meshKey, "Min_Y_Bound")) / numY
        gridArea = gridWidth * gridHeight
        'the start point of mesh coordinate 
        gridX0 = lt.DbGet(meshKey, "Min_X_Bound") + 0.5 * gridWidth
        gridY0 = lt.DbGet(meshKey, "Max_Y_Bound") - 0.5 * gridHeight
        'prepare position index and solid angle arrays
        ReDim posIdArray(numY - 1, numX - 1)
        ReDim strArray(numY - 1, numX - 1)
        minStr = Double.MaxValue
        maxStr = Double.MinValue
        For i = 0 To numY - 1
            For j = 0 To numX - 1
                gridX = gridX0 + gridWidth * j
                gridY = gridY0 - gridHeight * i
                s = Sqrt(workDist ^ 2 + (gridX ^ 2 + gridY ^ 2))
                posIdArray(i, j) = posInd(gridX, gridY, workDist)
                strArray(i, j) = gridStr(gridTheta(gridX, gridY, workDist), gridArea, s, workDist)
                If strArray(i, j) >= maxStr Then
                    maxStr = strArray(i, j)
                End If
                If strArray(i, j) <= minStr Then
                    minStr = strArray(i, j)
                End If
            Next
        Next
        Call TwoDArrayToCSV(posIdArray)
    End Sub
    Private Sub meshList_DropDownClosed(sender As Object, e As EventArgs) Handles meshList.DropDownClosed
        Call initMesh()
    End Sub
    Private Sub calcDimension()
        'calculate dummy plane dimensions according to PRR/Lum. mesh settings
        Dim fov, arW, arH As Double
        Dim imgWidth, imgHeight, workDist As Double
        workDist = CDbl(Me.workDist.Text)
        arW = CDbl(Me.arW.Text)
        arH = CDbl(Me.arH.Text)
        fov = CDbl(Me.fov.Text)
        imgHeight = 2 * workDist * Math.Tan((fov * Math.PI / 180) / 2)
        imgWidth = imgHeight * arW / arH
        Me.imgHeight.Text = imgHeight
        Me.imgWidth.Text = imgWidth
    End Sub
    Private Sub calcMinMeshStr()
        'calculate minimal solid angle on the mesh (the one on upper left corner)
        Dim gridWidth, gridHeight As Double
        Dim meshWidth, meshHeight As Double
        Dim r, d, s As Double
        Dim minMeshStr As Double
        meshWidth = imgWidth.Text
        meshHeight = imgHeight.Text
        d = workDist.Text
        gridWidth = meshWidth / xNum.Text
        gridHeight = meshHeight / yNum.Text
        r = Sqrt((meshWidth * 0.5 - gridWidth * 0.5) ^ 2 + (meshHeight * 0.5 - gridHeight * 0.5) ^ 2)
        s = Sqrt(r ^ 2 + d ^ 2)
        minMeshStr = (gridHeight * gridWidth) * d / s ^ 3
    End Sub
    Private Sub chkIntInput(sender As Object)
        Dim input As Integer
        If Not Integer.TryParse(DirectCast(sender, TextBox).Text, input) Then
            MsgBox("Integer value is required in this field!")
            DirectCast(sender, TextBox).Undo()
        End If
    End Sub
    Private Sub chkNumInput(sender As Object)
        Dim input As Double
        If Not Double.TryParse(DirectCast(sender, TextBox).Text, input) Then
            MsgBox("Numerical value is required in this field!")
            DirectCast(sender, TextBox).Undo()
        End If
    End Sub
    Private Sub workDist_KeyUp(sender As Object, e As KeyEventArgs) Handles workDist.KeyUp
        Call chkNumInput(sender)
        Call calcDimension()
        Call calcMinMeshStr()
    End Sub

    Private Sub fov_KeyUp(sender As Object, e As KeyEventArgs) Handles fov.KeyUp
        Call chkNumInput(sender)
        Call calcDimension()
        Call calcMinMeshStr()
    End Sub

    Private Sub apSize_KeyUp(sender As Object, e As KeyEventArgs) Handles apSize.KeyUp
        Call chkNumInput(sender)
    End Sub

    Private Sub lookD_KeyUp(sender As Object, e As KeyEventArgs) Handles lookD.KeyUp
        Call chkNumInput(sender)
    End Sub

    Private Sub xNum_KeyUp(sender As Object, e As KeyEventArgs) Handles xNum.KeyUp
        Call chkIntInput(sender)
        Call calcMinMeshStr()
    End Sub

    Private Sub yNum_KeyUp(sender As Object, e As KeyEventArgs) Handles yNum.KeyUp
        Call chkIntInput(sender)
        Call calcMinMeshStr()
    End Sub
    Private Sub arH_KeyUp(sender As Object, e As KeyEventArgs) Handles arH.KeyUp
        Call chkNumInput(sender)
        Call calcDimension()
    End Sub

    Private Sub arW_KeyUp(sender As Object, e As KeyEventArgs) Handles arW.KeyUp
        Call chkNumInput(sender)
        Call calcDimension()
    End Sub
    Private Sub createCamera(fov As Double, lookDist As Double, arW As Double, arH As Double)
        'create PRR camera
        lt.Cmd("\V3D")
        lt.Cmd("PlaceCamera " + js.LTCoord3(0, 0, 0) + " " + js.LTCoord3(0, 0, 1)) 'create a camera at origin
        newCamKey = "LENS_MANAGER[1].STUDIO_MANAGER[Studio_Manager].DATABASE[Camera_List].ENTITY[@Last]"
        newCamName = lt.DbGet(newCamKey, "Name")
        lt.DbSet(newCamKey, "Field_Of_View", fov)
        lt.DbSet(newCamKey, "Look_Distance", lookDist)
        lt.DbSet(newCamKey, "Aspect_Width", arW)
        lt.DbSet(newCamKey, "Aspect_Height", arH)
    End Sub
    Private Sub createImgPlane(width As Double, height As Double, xNum As Integer, yNum As Integer, workDist As Double, apSize As Double)
        'create image plane for pin hole imaging
        Dim newReceiverKey As String
        Dim newLumMeshKey As String
        stat = js.MakeDummyPlane(lt, 0, 0, -workDist, 0, 0, 1, W:=width, H:=height, ApertureType:="Rectangular") 'create dummy plane
        Debug.Print("dummy plane creating status: " + Str(stat))
        newDummyPlaneKey = "LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"
        newDummyPlaneName = lt.DbGet(newDummyPlaneKey, "Name")
        js.MakeReceiver(lt, EntityName:=newDummyPlaneName, ReceiverName:="Pinhole Image") 'create receiver
        'receiver settings
        newReceiverKey = "LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[@Last]"
        lt.DbSet(newReceiverKey + ".BACKWARD[Backward_Simulation]", "Has_Spatial_Luminance", "Yes")
        lt.DbSet(newReceiverKey + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "Meter_Collection_Mode", "Fixed Aperture")
        lt.DbSet(newReceiverKey + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "VirtualAperture", "Yes")
        lt.DbSet(newReceiverKey + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "Distance", workDist)
        lt.DbSet(newReceiverKey + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "Disk_Radius", apSize)
        'luminance mesh settings
        newLumMeshKey = newReceiverKey + ".BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]"
        lt.DbSet(newLumMeshKey, "X_Dimension", xNum)
        lt.DbSet(newLumMeshKey, "Y_Dimension", yNum)
        lt.DbSet(newLumMeshKey, "Name", "PinHoleLumMap")
    End Sub
    Private Sub initSettings()
        If Me.IsInitialized Then
            'Luminance mesh settings
            Me.xNum.Text = 320 'number of bins in x-direction
            Me.yNum.Text = 200 'number of bins in y-direction
            Me.workDist.Text = 20 'working distance
            Me.apSize.Text = 0.02 'aperture size
            Me.imgHeight.Text = 18.6523
            Me.imgWidth.Text = 24.8697
            'PRR camera settings
            Me.fov.Text = 50 'FoV
            Me.lookD.Text = 10 'looking distance
            Me.arW.Text = 4 'aspect ratio (width)
            Me.arH.Text = 3 'aspect raito (height)
        End If
    End Sub
    Private Sub reset_Click(sender As Object, e As RoutedEventArgs) Handles reset.Click
        Call initSettings()
    End Sub
    Private Sub create_Click(sender As Object, e As RoutedEventArgs) Handles create.Click
        Dim fov, lookD, arW, arH As Double
        Dim imgWidth, imgHeight, workDist, apSize As Double
        Dim xNum, yNum As Integer
        xNum = CInt(Me.xNum.Text)
        yNum = CInt(Me.yNum.Text)
        workDist = CDbl(Me.workDist.Text)
        apSize = CDbl(Me.apSize.Text)
        arW = CDbl(Me.arW.Text)
        arH = CDbl(Me.arH.Text)
        fov = CDbl(Me.fov.Text)
        lookD = CDbl(Me.lookD.Text)
        imgWidth = CDbl(Me.imgWidth.Text)
        imgHeight = CDbl(Me.imgHeight.Text)
        lt.SetOption("View Update", 0)
        lt.Begin()
        Call createCamera(fov, lookD, arW, arH)
        Call createImgPlane(imgWidth, imgHeight, xNum, yNum, workDist, apSize)
        lt.Cmd("Select " + newCamName)
        lt.Cmd("More " + newDummyPlaneName)
        lt.Cmd("Group")
        lt.SetOption("View Update", 1)
        newGroupKey = "LENS_MANAGER[1].COMPONENTS[Components].GROUP[@Last]"
        newGroupName = lt.DbGet(newGroupKey, "Name")
        lt.DbSet(newGroupKey, "Name", "PinHoleCam")
        lt.End()
    End Sub
    Sub TwoDArrayToCSV(ByVal DataArray(,) As Double)
        Dim strTmp As String = ""
        Dim ofile As String = ""

        svDialog("csv|*.csv", "save as...", ofile)
        Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(ofile)

        For i As Int32 = DataArray.GetLowerBound(0) To DataArray.GetUpperBound(0)
            For j As Int32 = DataArray.GetLowerBound(1) To DataArray.GetUpperBound(1)
                strTmp += Str(DataArray(i, j)) + ","
            Next
            sw.WriteLine(strTmp)
            strTmp = ""
        Next
        sw.Flush()
        sw.Close()
    End Sub

    Sub svDialog(ByVal infilter As String, ByVal dtitle As String, ByRef outfile As String)
        Dim openFileDialog1 As New Microsoft.Win32.SaveFileDialog()
        With openFileDialog1
            .Filter = infilter
            .FilterIndex = 1
            .Title = dtitle
            .DefaultExt = Strings.Right(infilter, 3)
            .ShowDialog()
            outfile = openFileDialog1.FileName
            .RestoreDirectory = True
        End With
    End Sub
End Class

