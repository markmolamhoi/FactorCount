Imports System.IO
Imports System.Text.Encoding

Public Class FactorCount
    Public miStartPercentList As New ArrayList
    Public miInputList As New ArrayList
    Public mi100Percent As Integer = 100
    Public miFactor As Integer
    Public miFactorUpperLimit As Integer
    Public miFactorLowerLimit As Integer
    Public miErrorPercent As Double = 0
    Public miMinimumnInput As Integer = 20000
    Public miTrueRecultCounter As Integer = 0
    Public miFalseRecultCounter As Integer = 0
    Public miLoopCounter As Integer = 0
    Public msResultList As New ArrayList
    Public msResultSummary As New ArrayList

    Public mbLoading As Boolean = False

    Dim mdStartTime As Date
    Dim mdEndTime As Date

    Private Function bValid() As Boolean
        If bIsNumber(Me.txtFactor.Text) Then

        Else
            MsgBox("Please input correct Factor.")
            Return False
        End If

        If Me.ListBox1.Items.Count = 0 Then
            MsgBox("Please input Sample.")
            Return False
        End If
        Return True
    End Function

    Private Function bIsNumber(ByVal pObject As Object) As Boolean
        Try
            Return IsNumeric(pObject)
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Call AddNumber()
    End Sub

    Private Sub AddNumber()
        If bIsNumber(Me.txtSample.Text) Then
            Me.ListBox1.Items.Add(Me.txtSample.Text)
            Me.txtSample.Text = ""
            Me.txtSample.Focus()
        Else
            MsgBox("Please input number.")
            Me.txtSample.Focus()
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.ListBox1.Items.Clear()
        Me.Timer1.Stop()
    End Sub

    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click

        If bValid() = True Then

            Call Me.GenerateData()
        End If
    End Sub

    Private Sub btnDefaultGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefaultGenerate.Click

        Me.ListBox1.Items.Add(200)
        Me.ListBox1.Items.Add(135)
        Me.ListBox1.Items.Add(150)
        Me.ListBox1.Items.Add(210)
        Me.ListBox1.Items.Add(115)
        Me.txtFactor.Text = 125
        Call GenerateData()

    End Sub

#Region "Initial Setting"
    Private Function iFindMinimumnInput() As Integer
        Dim liMini As Integer = 20000
        For Each liInput As Integer In miInputList
            If liMini > liInput Then
                liMini = liInput
            End If
        Next

        Return liMini
    End Function
    Private Sub BuildStartPercentList()
        For Each liInput As Integer In miInputList
            miStartPercentList.Add(iMinimizedStartPercent(liInput))
        Next
    End Sub
    Public Function iMinimizedStartPercent(ByVal piInteger As Integer) As Integer
        If piInteger <> miMinimumnInput Then
            Return ((miFactor - miMinimumnInput) / (piInteger - miMinimumnInput)) * 100
        Else
            Return 100
        End If

    End Function
    Private Sub BuildInputList()
        For Each liItem In Me.ListBox1.Items
            miInputList.Add(CInt(liItem.ToString()))
        Next
        miInputList.Sort()
        miInputList.Reverse()

    End Sub

    Private Sub GenerateData()
        ' Minimize the Count 
        miStartPercentList.Clear()
        miInputList.Clear()
        msResultList.Clear()
        msResultSummary.Clear()
        Me.lblResult.Text = ""
        miFactor = Me.txtFactor.Text
        miTrueRecultCounter = 0
        miFalseRecultCounter = 0
        miLoopCounter = 0


        miFactorUpperLimit = miFactor * (mi100Percent + miErrorPercent)
        miFactorLowerLimit = miFactor * (mi100Percent - miErrorPercent)

        Call BuildInputList()

        miMinimumnInput = iFindMinimumnInput()

        Call BuildStartPercentList()

        Dim lsCoumnHeader As String = sColumnHeader()

        Me.msResultList.Add("Sample:")
        Me.msResultList.Add(lsCoumnHeader)
        Me.msResultList.Add("Factor:")
        Me.msResultList.Add(Me.txtFactor.Text)
        Me.msResultList.Add("Percentage:")
        mdStartTime = Now
        mbLoading = True
        Me.Timer1.Start()

        Call nsThread.ParallelFunctionWithOutInvoke(AddressOf sThread)
    End Sub
    Private Sub sThread()

        Call sFinding100Combination(0, 0, "")

        Call WriteArraylistToFile("Result.txt", Me.msResultList)
        '  Call WriteArraylistToFile("ResultSummary.txt", Me.msResultSummary)

        mbLoading = False
        mdEndTime = Now
        Process.Start(".")
    End Sub
    Private Function sColumnHeader() As String
        Dim lsText As String = ""
        For Each liItem In Me.miInputList
            lsText += liItem & " "
        Next
        Return lsText
    End Function
    Private Sub AddResult(ByVal psString As String)
        Me.lblResult.Text = psString

    End Sub

#End Region

#Region "sFinding100Combination"

    Private Sub sFinding100Combination(ByVal piPreviousPercent As Integer, _
                    ByVal piListPosition As Integer, _
                    ByVal psPrintString As String)
        Dim liCurrentPercent As Integer = miStartPercentList(piListPosition)
        While liCurrentPercent >= 0
            Dim liPercentSum As Integer = piPreviousPercent + liCurrentPercent
            Dim lsPrintString As String = psPrintString & liCurrentPercent & " "
            miLoopCounter += 1
            Select Case True
                Case liPercentSum > mi100Percent
                    ' Debug.Print("Update minimumn Percent outside while. Original = " & liCurrentPercent)
                    liCurrentPercent = mi100Percent - piPreviousPercent + 1
                    ' Debug.Print("Update minimumn Percent outside while. Updated = " & liCurrentPercent)

                Case liPercentSum < mi100Percent
                    If piListPosition < miStartPercentList.Count - 1 Then

                        ' If the Parent Q greater than Factor, then skip the rest.
                        If bCheckLimit(lsPrintString, liPercentSum) = False Then
                            If bUpdateCurrentPercent(psPrintString, piListPosition, liCurrentPercent, piPreviousPercent) = True Then

                            Else
                                '  Debug.Print("Exit Loop Negative Percent = " & lsPrintString)
                                Exit Sub
                            End If
                        Else
                            sFinding100Combination(liPercentSum, piListPosition + 1, lsPrintString)
                        End If

                    Else

                        '  Debug.Print("liPercent < mi100Percent - Loop all = " & lsPrintString)
                        Exit Sub
                    End If

                Case liPercentSum = mi100Percent
                    bCheck100Result(lsPrintString)

            End Select
            liCurrentPercent -= 1
        End While

    End Sub

    Private Function bUpdateCurrentPercent(ByVal psPrintString As String, _
                    ByVal piListPosition As Integer, _
                    ByRef liCurrentPercent As Integer, _
                    ByRef piPreviousPercent As Integer) As Boolean

        Dim liResultFactor As Double = dCalculateFactor(psPrintString)
        Dim liMinimunFactor As Double = miMinimumnInput * (mi100Percent - piPreviousPercent)
        Dim liParent As Double = miFactorUpperLimit - liResultFactor - liMinimunFactor
        Dim liSon As Double = miInputList(piListPosition) - miMinimumnInput

        ' Debug.Print("Update minimumn Percent inside while. Original = " & liCurrentPercent)
        Dim liCalculatePercent As Integer = liParent / liSon
        If liCalculatePercent = liCurrentPercent Then
            liCurrentPercent = liCalculatePercent
        Else
            liCurrentPercent = liCalculatePercent + 1
        End If
        ' Debug.Print("Update minimumn Percent inside while. Updated = " & liCurrentPercent)
        If liCurrentPercent >= 0 Then

            Return True
        Else
            Return False
        End If
    End Function

    Private Function dCalculateFactor(ByVal psPrintString As String) As Double
        ' 計算現有的 factor 值.
        Dim lList As String() = psPrintString.Split(" ")
        Dim liResultFactor As Double
        For i As Integer = 0 To lList.Length - 2
            liResultFactor += CDbl(lList(i)) * miInputList(i)
        Next
        Return liResultFactor
    End Function

    Private Function bCheck100Result(ByVal psPrintString As String) As Boolean
        ' 可以使用誤差
        Dim liResultFactor As Double = dCalculateFactor(psPrintString)
        Dim lbResult As Boolean = False
        If miFactorLowerLimit <= liResultFactor And liResultFactor <= miFactorUpperLimit Then
            lbResult = True
            ' Debug.Print("100% = " & psPrintString & " True Result : = " & liResultFactor.ToString)
            msResultList.Add(psPrintString)
            miTrueRecultCounter += 1
        Else
            ' Debug.Print("100% = " & psPrintString & " False Result : = " & liResultFactor.ToString)
            lbResult = False
            miFalseRecultCounter += 1
        End If
        Return lbResult
    End Function

    Private Function bCheckLimit(ByVal psPrintString As String, ByVal psPercentSum As Integer) As Boolean
        Dim liResultFactor As Double = dCalculateFactor(psPrintString)
        Dim liRemainFactor As Double = (mi100Percent - psPercentSum) * miMinimumnInput
        liResultFactor = liResultFactor + liRemainFactor

        Dim lbResult As Boolean = False
        If liResultFactor <= miFactorUpperLimit Then
            ' Debug.Print("within Limit : = " & psPrintString & " Factor : = " & liResultFactor.ToString)
            lbResult = True

        Else
            ' Debug.Print("Over Limit : = " & psPrintString & " Factor : = " & liResultFactor.ToString)
            lbResult = False

        End If
        Return lbResult
    End Function
#End Region

    Public Function WriteArraylistToFile(ByVal psDirPath As String, ByVal pArrayList As ArrayList) As ArrayList
        Dim compare_arraylist As ArrayList = New ArrayList

        If File.Exists(psDirPath) Then
            File.Delete(psDirPath)
        End If

        Dim fs As FileStream = New FileStream(psDirPath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim w As StreamWriter = New StreamWriter(fs)
        w.BaseStream.Seek(0, SeekOrigin.Begin)
        Dim aaa As String
        For Each aaa In pArrayList
            If compare_arraylist.Contains(aaa) Then
            Else
                compare_arraylist.Add(aaa)
                w.WriteLine(aaa, Unicode)
            End If

        Next
        w.Flush() ' update underlying file
        w.Close()
        w.Dispose()

        Return pArrayList
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim totalTime As TimeSpan
        If mbLoading Then
            totalTime = Now.Subtract(mdStartTime)
        Else
            totalTime = mdEndTime.Subtract(mdStartTime)
        End If
        Dim lsString As String = "Result.txt is the Data Result." & vbNewLine & _
            "No Of True Result:" & miTrueRecultCounter & vbNewLine & _
             "Looping Counter:" & miLoopCounter & vbNewLine & _
              "No Of False Result:" & miFalseRecultCounter & vbNewLine & _
                " Time take : " & totalTime.Duration.ToString

        Call AddResult(lsString)

    End Sub

    Private Sub txtSample_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSample.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call AddNumber()
        End If
    End Sub

End Class
