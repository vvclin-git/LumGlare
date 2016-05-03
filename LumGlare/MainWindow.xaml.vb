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


    Private Sub test_Click(sender As Object, e As RoutedEventArgs) Handles test.Click
        Dim receiverKey, meshKey As String
        Dim numSample As Long
        Dim lumArray(,) As Double
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
            ltStat = lt.GetMeshData(meshKey, lumArray, "CellValue")
            ugrTmp = 0
            For i = 0 To numY - 1
                For j = 0 To numX - 1
                    ugrTmp += (lumArray(j, i)) ^ 2 * strArray(j, i) / (posIdArray(j, i)) ^ 2
                Next
            Next
            'Debug.Print(ugrTmp)
            UGR = 8 * Log10(0.25 / backLum * ugrTmp)
            Me.ugrOut.Text = String.Format("{0:0.0}", UGR)
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
        'Debug.Print("theta:" + Str(theta))
        'Debug.Print("tau:" + Str(tau))
        If theta > 0 Then
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
        ReDim posIdArray(numX - 1, numY - 1)
        ReDim strArray(numX - 1, numY - 1)
        minStr = Double.MaxValue
        maxStr = Double.MinValue
        For i = 0 To numY - 1
            For j = 0 To numX - 1
                gridX = gridX0 + gridWidth * i
                gridY = gridY0 - gridHeight * j
                s = Sqrt(workDist ^ 2 + (gridX ^ 2 + gridY ^ 2))
                posIdArray(j, i) = posInd(gridX, gridY, workDist)
                strArray(j, i) = gridStr(gridTheta(gridX, gridY, workDist), gridArea, s, workDist)
                If strArray(j, i) >= maxStr Then
                    maxStr = strArray(j, i)
                End If
                If strArray(j, i) <= minStr Then
                    minStr = strArray(j, i)
                End If
            Next
        Next
        Me.minStrOut.Text = String.Format("{0:0.0##e+00}", minStr)
        If minStr < 0.0002 Then
            Me.strStat.Content = "sr (Mesh solid angle should >2e-05 !)"
        Else
            Me.strStat.Content = "sr"
        End If

    End Sub
    Private Sub meshList_DropDownClosed(sender As Object, e As EventArgs) Handles meshList.DropDownClosed
        Call initMesh()
    End Sub
End Class

