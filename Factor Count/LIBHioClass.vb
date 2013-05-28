Imports System.IO
Imports System.Threading

Public Module LIBHioClass
    Public Delegate Sub DSub0()
    Public Delegate Sub DSub1(ByVal pObject As Object)
    Public Delegate Sub DSub2(ByVal pObject As Object, ByRef pObject As Object)
    Public Delegate Sub DSub1Grid(ByRef pGrid As Windows.Forms.DataGridView)
    Public Const APP_NAME As String = "Factor Count"
    Public Sub ErrHandler(ByVal ErrMsg As String, ByVal ErrTitle As String)

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

            'Call Beep()

            MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTitle)

            Debug.Print(ErrMsg)
        Catch Err As Exception
            Call MsgBox(Err.Message, MsgBoxStyle.Critical, ErrTitle)
        End Try

    End Sub

End Module

#Region "nsThread"
Namespace nsThread
    Public Module mThread

        ''' <summary>
        ''' It will auto use the invoke when it is required, however the performance is not the best.
        ''' </summary>
        Public Function ParallelFunction(ByVal pForm As Form, ByVal pSub0 As DSub0) As Thread
            ParallelFunction = Nothing
            Try
                Dim ClsThreadHelper As New ClassThreadHelper
                ClsThreadHelper.SetForm(pForm)
                ClsThreadHelper.SetSub0(pSub0)
                ClsThreadHelper.StartThreadFunction()
                ParallelFunction = ClsThreadHelper.mThread
            Catch Err As Exception

            End Try
        End Function

        ''' <summary>
        ''' same as above but using DSub1.
        ''' </summary>
        Public Function ParallelFunction(ByVal pForm As Form, ByVal pSub1 As DSub1, ByVal pobject As Object) As Thread
            ParallelFunction = Nothing

            Try
                Dim ClsThreadHelper As New ClassThreadHelper
                ClsThreadHelper.SetForm(pForm)
                ClsThreadHelper.SetSub1(pSub1)
                ClsThreadHelper.SetParameter(pobject)
                ClsThreadHelper.StartThreadFunction()
                ParallelFunction = ClsThreadHelper.mThread
            Catch Err As Exception

            End Try
        End Function

        ''' <summary>
        ''' use to build a new thread when the new thread does not interact with main thread.
        ''' </summary>
        Public Function ParallelFunctionWithOutInvoke(ByVal pSub0 As DSub0) As Thread
            ParallelFunctionWithOutInvoke = Nothing
            Try
                Dim ClsThreadHelperBasic As New ClassThreadHelperBasic
                ClsThreadHelperBasic.SetSub0(pSub0)
                ClsThreadHelperBasic.StartThreadFunctionWithOutInvoke()
                ParallelFunctionWithOutInvoke = ClsThreadHelperBasic.mThread
            Catch Err As Exception

            End Try
        End Function

        ''' <summary>
        ''' same as above but using DSub1.
        ''' "I also want to combine 2 functions into 1, however I can not find a solution. by Hoi"
        ''' </summary>
        Public Function ParallelFunctionWithOutInvoke(ByVal pSub1 As DSub1, ByVal pobject As Object) As Thread
            ParallelFunctionWithOutInvoke = Nothing
            Try
                Dim ClsThreadHelperBasic As New ClassThreadHelperBasic
                ClsThreadHelperBasic.SetSub1(pSub1)
                ClsThreadHelperBasic.SetParameter(pobject)
                ClsThreadHelperBasic.StartThreadFunctionWithOutInvoke()
                ParallelFunctionWithOutInvoke = ClsThreadHelperBasic.mThread

            Catch Err As Exception

            End Try
        End Function

        ''' <summary>
        ''' When the new thread wants to interact with the main thread, it required invokeFunction.
        ''' </summary>
        Public Sub InvokeFunction(ByVal pForm As Form, ByVal pSub0 As DSub0)
            Try
                Dim ClsThreadHelper As New ClassThreadHelper
                ClsThreadHelper.SetForm(pForm)
                ClsThreadHelper.SetSub0(pSub0)
                ClsThreadHelper.InvokeFunction()

            Catch Err As Exception

            End Try
        End Sub

        ''' <summary>
        ''' same as above but using DSub1.
        ''' </summary>
        Public Sub InvokeFunction(ByVal pForm As Form, ByVal pSub1 As DSub1, ByVal pParameter As Object)
            Try
                Dim ClsThreadHelper As New ClassThreadHelper
                ClsThreadHelper.SetForm(pForm)
                ClsThreadHelper.SetSub1(pSub1)
                ClsThreadHelper.SetParameter(pParameter)
                ClsThreadHelper.InvokeFunction()

            Catch Err As Exception

            End Try
        End Sub
    End Module
End Namespace
#End Region

#Region " ClassThreadHelper "
' This class is rename from ClassFunctionHelper to ClassThreadHelper
''' <summary>
''' used to handle invokeRequired thread.
''' </summary>
Public Class ClassThreadHelper
    Inherits ClassThreadHelperBasic

    Delegate Sub CallbackFunction()

    Private mForm As Form

    ''' <summary>
    ''' used to check the function required invoke or not. mForm represents main thread.
    ''' </summary>
    Public Sub SetForm(ByVal pForm As Form)
        Try
            Me.mForm = pForm
        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    Public Sub InvokeFunction()
        Try

            If mForm.InvokeRequired Then
                Dim CallBackForTest As New CallbackFunction(AddressOf InvokeFunction)
                mForm.Invoke(CallBackForTest)
            Else

                MyBase.DSubThread()

            End If
        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    Public Sub StartThreadFunction()
        Try

            MyBase.mThread = New Thread(New ThreadStart(AddressOf InvokeFunction))
            MyBase.mThread.Start()
        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

End Class

''' <summary>
''' used to handle simple thread. (Without Invoke)
''' </summary>
Public Class ClassThreadHelperBasic
    Protected mDSub0 As DSub0
    Protected mDSub1 As DSub1
    Protected mObject As Object
    Public mThread As Thread = Nothing

    ''' <summary>
    ''' Sets the sub1. you should also call SetParameter
    ''' </summary>
    Public Sub SetSub1(ByVal pDSub1 As DSub1)
        Try

            Me.mDSub1 = pDSub1

        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    Public Sub SetParameter(ByVal pObject As Object)
        Try

            mObject = pObject

        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    ''' <summary>
    ''' Sets the sub0. No need parameter
    ''' </summary>
    Public Sub SetSub0(ByVal pDSub0 As DSub0)
        Try

            Me.mDSub0 = pDSub0


        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    Public Sub StartThreadFunctionWithOutInvoke()
        Try

            If Me.mDSub0 Is Nothing Then
                ' cannot directly call DSubThread, as it will cause thread problem.
                mThread = New Thread(New ParameterizedThreadStart(AddressOf DSub1Thread))
                mThread.Start(Me.mObject)
            Else
                mThread = New Thread(New ThreadStart(AddressOf DSubThread))
                mThread.Start()
            End If

        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    ''' <summary>
    ''' In this class, either mDSub0 or mDSub1 is an instance.
    ''' </summary>
    Protected Sub DSubThread()
        Try
            If Me.mDSub0 Is Nothing Then
                Me.mDSub1(Me.mObject)
            Else
                Me.mDSub0()
            End If

        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

    Private Sub DSub1Thread(ByVal pObject As Object)
        Try

            Me.mDSub1(Me.mObject)

        Catch Err As Exception
            Call ErrHandler(Err.Message, APP_NAME)
        End Try
    End Sub

End Class
#End Region
