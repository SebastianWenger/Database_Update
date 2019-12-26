Imports sapfewse

Public Class SAPGUI
    Private SapCOM As Object = Nothing
    Private SAPApp As GuiApplication = Nothing
    Private SapConn As GuiConnection = Nothing
    Private SapSession As GuiSession = Nothing
    Public SaveDirectory As String = AppDomain.CurrentDomain.BaseDirectory
    Private SAPSaveName As String = Nothing
    Private FileExt As Integer = 31
    Private SalesOrg = New String() {"US01", "US02", "CA01"}


    Public ReadOnly Property StatusBarMessageType() As Char

        Get
            StatusBarMessageType = session.FindById("wnd[0]/sbar").MessageType
        End Get

    End Property

    Public ReadOnly Property StatusBarText() As String

        Get
            StatusBarText = session.FindById("wnd[0]/sbar").text
        End Get

    End Property

    Public ReadOnly Property session As GuiSession

        Get
            session = SapSession
        End Get

    End Property

    Public ReadOnly Property SaveName() As String

        Get
            SaveName = SAPSaveName
        End Get

    End Property


    Public Sub New(session_num As String)

        SapCOM = GetObject("SAPGUISERVER")
        SAPApp = SapCOM.GetScriptingEngine
        SapConn = SAPApp.Children(0)
        SapSession = SapConn.Children(0 + session_num)

        SAPSaveName = "ZSD_CONT_LIST_" + session_num + ".XLSX"

    End Sub

    Public Function FindByText(Search As String) As Object

        FindByText = Nothing
        Try
            For Each Children As Object In session.FindById("wnd[0]/usr/").children
                If Children.Text.ToString.Trim = Search.Trim Then
                    FindByText = Children
                    Exit For
                End If
            Next
        Catch ex As Exception
        End Try

    End Function

    Public Function ActiveWindow() As GuiFrameWindow

        ActiveWindow = session.ActiveWindow

    End Function

    Public Function UserArea() As GuiUserArea

        UserArea = ActiveWindow.FindById("usr")

    End Function

    Public Sub SendVKey(ByVal Code As Integer)

        session.ActiveWindow.SendVKey(Code)

    End Sub

    Public Sub StartTransaction(ByVal Code As String)

        session.StartTransaction(Code)

    End Sub

    Public Sub ClearAllGUITextField()

        For Each text_field As Object In UserArea().Children
            Try
                If text_field.Text <> "" Then text_field.Text = ""
            Catch
            End Try
        Next

    End Sub

    Public Sub Set_SalesOrg()
        multiple_selection("VKORG")
        fill_multiple_selection(SalesOrg)
        SendVKey(0)
        SendVKey(8)
    End Sub

    Public Sub Fill_LowHigh(name As String, low_value As String, high_value As String)
        Fill_TextField_Name(name + "-LOW", low_value)
        Fill_TextField_Name(name + "-HIGH", high_value)
    End Sub


    Private Sub btn_press(name As String)

        Dim btn As GuiButton = UserArea.FindByName(name, "GuiButton")
        btn.Press()

    End Sub

    Private Sub multiple_selection(name As String)

        btn_press("%_" + name + "_%_APP_%-VALU_PUSH")

    End Sub

    Private Sub fill_multiple_selection(data As String())

        Dim row_count As Long = data.LongCount - 1
        Dim tbl As GuiTableControl = UserArea.FindByName("SAPLALDBSINGLE", "GuiTableControl")

        For row As Long = 0 To row_count
            tbl.FindById("ctxtRSCSEL_255-SLOW_I[1," & row & "]").Text = data(row)
        Next

    End Sub


    Public Sub Fill_TextField_Name(name As String, input As String)


        If UserArea.FindByName(name, "GuicTextField") Is Nothing Then
            Dim txtfield As GuiTextField = UserArea.FindByName(name, "GuiTextField")
            txtfield.Text = input
        Else
            Dim txtfield As GuiCTextField = UserArea.FindByName(name, "GuiCTextField")
            txtfield.Text = input
        End If


    End Sub

    Public Sub Fill_TextField_ID(id As String, input As String)


        If UserArea.FindById(id) Is Nothing Then
            Dim txtfield As GuiTextField = UserArea.FindById(id)
            txtfield.Text = input
        Else
            Dim txtfield As GuiCTextField = UserArea.FindById(id)
            txtfield.Text = input
        End If


    End Sub

    Public Sub Save_Report()

        If IO.File.Exists(SaveDirectory + SaveName) Then My.Computer.FileSystem.DeleteFile(SaveDirectory + SaveName)

        ActiveWindow.SendVKey(8)

        ActiveWindow.SendVKey(21)

        Dim combobox As GuiComboBox = UserArea.FindByName("G_LISTBOX", "GuiComboBox")
        combobox.Key = FileExt

        ActiveWindow.SendVKey(0)

        Fill_TextField_ID("ctxtDY_PATH", SaveDirectory)

        Fill_TextField_ID("ctxtDY_FILENAME", SaveName)

        ActiveWindow.SendVKey(0)

    End Sub

    Sub Workbook_Close()

        Dim p As New Process

        p.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "\SAP_Workbook_Close.vbs"
        p.StartInfo.Arguments = SaveName
        p.Start()
        p.WaitForExit()
        p.Close()

    End Sub

End Class
