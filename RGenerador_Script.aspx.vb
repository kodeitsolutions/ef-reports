'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "RGenerador_Script"
'-------------------------------------------------------------------------------------------'
Partial Class RGenerador_Script
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try



            Dim loComandoSeleccionar As New StringBuilder()




            loComandoSeleccionar.AppendLine("USE aap;")
            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("            CAST(OBJECT_NAME (tablas.object_id) as varchar) as tabla,      ")
            loComandoSeleccionar.AppendLine("            CAST(columnas.name as varchar)as columna,    ")
            loComandoSeleccionar.AppendLine("            CAST(tipos.name as varchar)as tipo,    ")
            loComandoSeleccionar.AppendLine("            CAST(defs.definition as varchar) as valor_default,  ")
            loComandoSeleccionar.AppendLine("            CASE ISNULL(cast(defs.definition as varchar),'1')  ")
            loComandoSeleccionar.AppendLine("            WHEN '1' THEN '1'  ")
            loComandoSeleccionar.AppendLine("            ELSE '2'  ")
            loComandoSeleccionar.AppendLine("            END AS valor_default2  ")

            loComandoSeleccionar.AppendLine("FROM       ")
            loComandoSeleccionar.AppendLine("           sys.objects AS tablas    ")
            loComandoSeleccionar.AppendLine("           JOIN sys.columns as columnas on tablas.object_id = columnas.object_id      ")
            loComandoSeleccionar.AppendLine("           JOIN sys.types as tipos on tipos.user_type_id = columnas.user_type_id      ")
            loComandoSeleccionar.AppendLine("           LEFT join sys.default_constraints as defs  on defs.parent_object_id  = tablas.object_id    ")
            loComandoSeleccionar.AppendLine("           AND  defs.parent_column_id = columnas.column_id  ")
            loComandoSeleccionar.AppendLine("WHERE      tablas.type = 'U'    ")
            loComandoSeleccionar.AppendLine("ORDER BY   tablas.name; ")



            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")




            Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
            Dim loFila As DataRow

            loComandoSeleccionar = New StringBuilder()    '**************************************

            For lnNumeroFila As Integer = 0 To lnTotalFilas - 1

                loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)

                Dim lcValorDefault As String = Trim(loFila("valor_default2"))

                If (lcValorDefault = 1) Then

                    Dim lcTipo As String = Trim(loFila("tipo"))

                    Select Case lcTipo


                        Case "int"

                            'loComandoSeleccionar.AppendLine(" ALTER" & Trim(loFila("tabla")) & " COLUMN" & Trim(loFila("columna")))'
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT (0) FOR " & Trim(loFila("columna")))

                        Case "char"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT space(1) FOR " & Trim(loFila("columna")))

                        Case "text"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT space(1) FOR " & Trim(loFila("columna")))

                        Case "decimal"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT 0 FOR " & Trim(loFila("columna")))

                        Case "double"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT 0 FOR " & Trim(loFila("columna")))

                        Case "bit"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT 0 FOR " & Trim(loFila("columna")))

                        Case "datetime"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT getdate() FOR " & Trim(loFila("columna")))

                        Case "float"
                            loComandoSeleccionar.AppendLine(" ALTER TABLE " & Trim(loFila("tabla")) & "  ADD DEFAULT 0 FOR " & Trim(loFila("columna")))




                        Case Else


                    End Select

                End If


            Next lnNumeroFila

            Response.Clear()

            Response.Write("<pre>")
            Response.Write(loComandoSeleccionar.ToString())
            Response.Write("</pre>")
            Response.Flush()
            Response.End()



        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try


    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub
End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'

