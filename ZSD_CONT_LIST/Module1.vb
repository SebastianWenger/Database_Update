

Module Module1

    Sub Main(ByVal args As String())

        Dim begin_date As String = args(0)
        Dim end_date As String = args(1)
        Dim refresh_database As Boolean = args(2)
        Dim session_num As Short = args(3)


        Dim sapgui = New SAPGUI(session_num)

        Try
            Do
                With sapgui
                    .StartTransaction("ZSD_CONT_LIST")
                    .ClearAllGUITextField()
                    .Set_SalesOrg()
                    .Fill_TextField_Name("VBTYP-LOW", "G")
                    .Fill_LowHigh("ERDAT_L", begin_date, end_date) 'Line Created
                    '.Fill_LowHigh("AEDAT_I", begin_date, end_date) 'Line Changed
                    .Save_Report()
                    .Workbook_Close()
                End With






                If refresh_database = True Then

                    Dim sqldatabase As New SQLDatabase

                    With sqldatabase
                        Dim dt As DataTable = .ImportExceltoDatatable(sapgui.SaveDirectory + sapgui.SaveName)
                        .AddDateColumn(dt)
                        .Insert(dt)
                        .RemoveDuplicates()
                    End With

                End If
            Loop

        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Console.ReadKey()
        End Try




    End Sub






End Module
