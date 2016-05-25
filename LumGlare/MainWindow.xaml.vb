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
    Dim debugMode As Boolean


    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        debugMode = False
        If Not debugMode Then
            lt = LTGetter.getLTAPIServer
            If IsNothing(lt) Then
                MsgBox("LightTools session not found!")
            Else
                Me.Title = "Luminance Map UGR Calculator (PID = " + Str(lt.GetServerID()) + ")"
                Call refreshMeshList()
            End If
        Else
            pID = 15940
            lt = ltLoc.GetLTAPI(pID)
            Me.Title = "Luminance Map UGR Calculator (PID = " + Str(lt.GetServerID()) + ")"
            Call refreshMeshList()
        End If
    End Sub

    Private Sub refreshMeshList()
        'retrieve a list of luminance/illuminance meshes in LT
        Dim receiverListKey, receiverKey As String
        Dim meshListKey, meshKey As String
        Dim meshName As String
        Dim statList1 As LTReturnCodeEnum
        Dim statList2 As LTReturnCodeEnum
        Dim statList3 As LTReturnCodeEnum
        'get receiver list
        receiverListKey = lt.DbList("LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager]", "SURFACE_RECEIVER")
        'loop through the receivers
        Do
            receiverKey = lt.ListNext(receiverListKey, statList1)
            'only deal with the receivers with keyword in their names
            'If lt.DbGet(receiverKey, "Name") Like "PinHole Image*" Then
            'for each receiver, return a list containing all meshes
            meshListKey = lt.DbList(receiverKey, "SPATIAL_LUMINANCE_MESH", statList3)
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
                'If lt.DbType(meshKey, "INTENSITY_MESH") = 0 Then
                'If lt.DbGet(meshKey, "Name") Like "PinHoleLumMap*" Then
                meshName = lt.DbGet(receiverKey, "Name") + " | " + lt.DbGet(meshKey, "Name")
                If Not meshKeyDict.ContainsKey(meshName) Then
                    'put mesh name-key mapping into dictionary
                    meshKeyDict.Add(meshName, meshKey)
                    'put mesh name-receiver key mapping into dictionary
                    receiverKeyDict.Add(meshName, receiverKey)
                    'add meshName to the combobox
                    Me.meshList.Items.Add(meshName)
                End If
                'End If
                'End If
            Loop
            'End If
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
            MsgBox("No ray data collected in the mesh!")
        Else
            ReDim lumArray(numX - 1, numY - 1)
            ReDim ugrArray(numY - 1, numX - 1)
            ltStat = lt.GetMeshData(meshKey, lumArray, "CellValue")
            ugrTmp = 0
            For i = 0 To numY - 1
                For j = 0 To numX - 1
                    'Debug.Print("calculating: " + Str(i) + ", " + Str(j))
                    ugrTmp += (lumArray(j, i)) ^ 2 * strArray(i, j) / (posIdArray(i, j)) ^ 2
                    ugrArray(i, j) = (lumArray(j, numY - 1 - i)) ^ 2 * strArray(i, j) / (posIdArray(i, j)) ^ 2
                Next
            Next
            'Debug.Print(ugrTmp)
            UGR = 8 * Log10(0.25 / backLum * ugrTmp)
            Me.ugrOut.Text = String.Format("{0:0.0}", UGR)
            'Call TwoDArrayToCSV(lumArray) 'dump luminance data into csv (for debugging)
            'Call TwoDArrayToCSV(ugrArray) 'dump calculated ugrTmp data into csv (for debugging)
        End If
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
        Dim meshWidth, meshHeight As Double
        Dim minMeshStr As Double
        Dim s, r As Double
        Dim numX, numY As Integer
        receiverKey = receiverKeyDict(Me.meshList.Text)
        meshKey = meshKeyDict(Me.meshList.Text)
        workDist = lt.DbGet(lt.DbKeyStr(receiverKey) + ".SPATIAL_LUM_METER[Spatial_Lum_Meter]", "Distance")
        'get grid number and dimension from mesh
        numX = lt.DbGet(meshKey, "X_Dimension")
        numY = lt.DbGet(meshKey, "Y_Dimension")
        meshWidth = lt.DbGet(meshKey, "Max_X_Bound") - lt.DbGet(meshKey, "Min_X_Bound")
        meshHeight = lt.DbGet(meshKey, "Max_Y_Bound") - lt.DbGet(meshKey, "Min_Y_Bound")
        gridWidth = (lt.DbGet(meshKey, "Max_X_Bound") - lt.DbGet(meshKey, "Min_X_Bound")) / numX
        gridHeight = (lt.DbGet(meshKey, "Max_Y_Bound") - lt.DbGet(meshKey, "Min_Y_Bound")) / numY
        gridArea = gridWidth * gridHeight
        'the start point of mesh coordinate 
        gridX0 = lt.DbGet(meshKey, "Min_X_Bound") + 0.5 * gridWidth
        gridY0 = lt.DbGet(meshKey, "Max_Y_Bound") - 0.5 * gridHeight
        'prepare position index and solid angle arrays
        ReDim posIdArray(numY - 1, numX - 1)
        ReDim strArray(numY - 1, numX - 1)
        For i = 0 To numY - 1
            For j = 0 To numX - 1
                gridX = gridX0 + gridWidth * j
                gridY = gridY0 - gridHeight * i
                s = Sqrt(workDist ^ 2 + (gridX ^ 2 + gridY ^ 2))
                posIdArray(i, j) = posInd(gridX, gridY, workDist)
                strArray(i, j) = gridStr(gridTheta(gridX, gridY, workDist), gridArea, s, workDist)
            Next
        Next
        'calculate minimal solid angle over mesh
        minMeshStr = calcMinMeshStr(meshWidth, meshHeight, numX, numY, workDist)
        minStrOut.Text = String.Format("{0:0.000e00}", minMeshStr)
        'Call TwoDArrayToCSV(posIdArray) 'for debugging
    End Sub
    Private Sub meshList_DropDownClosed(sender As Object, e As EventArgs) Handles meshList.DropDownClosed
        If (Me.meshList.Text <> "") Then
            Call initMesh()
        End If
    End Sub

    Private Function calcMinMeshStr(meshWidth As Double, meshHeight As Double, xNum As Integer, yNum As Integer, workDist As Double)
        'calculate minimal solid angle on the mesh (the one on upper left corner)
        Dim gridWidth, gridHeight As Double
        Dim r, d, s As Double
        Dim minMeshStr As Double
        d = workDist
        gridWidth = meshWidth / xNum
        gridHeight = meshHeight / yNum
        r = Sqrt((meshWidth * 0.5 - gridWidth * 0.5) ^ 2 + (meshHeight * 0.5 - gridHeight * 0.5) ^ 2)
        s = Sqrt(r ^ 2 + d ^ 2)
        minMeshStr = (gridHeight * gridWidth) * d / s ^ 3
        Return minMeshStr
    End Function

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

