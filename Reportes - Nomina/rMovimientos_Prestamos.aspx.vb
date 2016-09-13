'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMovimientos_Prestamos"
'-------------------------------------------------------------------------------------------'
Partial Class rMovimientos_Prestamos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT  Recibos.Documento	                                    AS Documento,	")
            loConsulta.AppendLine("        Recibos.fecha		                                    AS Fecha,		")
            loConsulta.AppendLine("        Recibos.Status		                                    AS Estatus,		")
            loConsulta.AppendLine("        Recibos.Cod_Tra		                                    AS Cod_Tra,		")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra                                     AS Nom_Tra,		")
            loConsulta.AppendLine("        Prestamo.Prestamo_Documento                              AS Prestamo_Documento,	")
            loConsulta.AppendLine("        (CASE WHEN Prestamo.Prestamo_Cuota = 0")
            loConsulta.AppendLine("            THEN 'Prestamo'")
            loConsulta.AppendLine("            ELSE 'Cuota #'")
            loConsulta.AppendLine("                +CAST(Prestamo.Prestamo_Cuota AS VARCHAR(20))")
            loConsulta.AppendLine("                +'/'")
            loConsulta.AppendLine("                +CAST(Prestamo.Prestamo_Cuotas AS VARCHAR(20))")
            loConsulta.AppendLine("        END)	                                                   AS Detalle,		")
            loConsulta.AppendLine("        (CASE WHEN Prestamo.Prestamo_Cuota = 0")
            loConsulta.AppendLine("            THEN Prestamo.Prestamo_Asignacion")
            loConsulta.AppendLine("            ELSE Prestamo.Prestamo_Vencimiento")
            loConsulta.AppendLine("        END)	                                                   AS Fecha_Prestamo,		")
            loConsulta.AppendLine("        Prestamo.Prestamo_Cuota                                 AS Cuota,  ")
            loConsulta.AppendLine("        Prestamo.Prestamo_Monto                                 AS Monto_Prestamo,    ")
            loConsulta.AppendLine("        Prestamo.Prestamo_Por_Int                               AS Prestamo_Por_Int,    ")
            loConsulta.AppendLine("        Prestamo.Prestamo_Mon_Int                               AS Prestamo_Mon_Int,    ")
            loConsulta.AppendLine("        Prestamo.Prestamo_Monto_Cuota                           AS Monto_Cuota,    ")
            loConsulta.AppendLine("        Prestamo.Prestamo_Motivo                                AS Cod_Mot,		")
            loConsulta.AppendLine("        COALESCE(Motivos.Nom_Mot, '')                           AS Nom_Mot,		")
            loConsulta.AppendLine("        Recibos.Comentario                                      AS Comentario	")
            loConsulta.AppendLine("FROM		Recibos ")
            loConsulta.AppendLine("    JOIN	Trabajadores ")
            loConsulta.AppendLine("        ON Trabajadores.cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("    JOIN(   SELECT  Renglones_Recibos.Documento     AS Prestamo_Recibo,")
            loConsulta.AppendLine("                    Prestamos.Documento             AS Prestamo_Documento,")
            loConsulta.AppendLine("                    Prestamos.Motivo                AS Prestamo_Motivo,")
            loConsulta.AppendLine("                    Prestamos.Fec_Asi               AS Prestamo_Asignacion,")
            loConsulta.AppendLine("                    CAST(Prestamos.Cuotas AS INT)   AS Prestamo_Cuotas,")
            loConsulta.AppendLine("                    Renglones_Recibos.Ren_Ori       AS Prestamo_Cuota,")
            loConsulta.AppendLine("                    Detalles_Prestamos.Fec_Ven      AS Prestamo_Vencimiento,")
            loConsulta.AppendLine("                    SUM(CASE ")
            loConsulta.AppendLine("                            WHEN Renglones_Recibos.Tipo = 'Asignacion'")
            loConsulta.AppendLine("                            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        END)                        AS Prestamo_Monto,")
            loConsulta.AppendLine("                    MAX(CASE ")
            loConsulta.AppendLine("                            WHEN Renglones_Recibos.Tipo = 'Asignacion'")
            loConsulta.AppendLine("                            THEN Prestamos.Por_Int")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        END)                        AS Prestamo_Por_Int,")
            loConsulta.AppendLine("                    MAX(CASE ")
            loConsulta.AppendLine("                            WHEN Renglones_Recibos.Tipo = 'Asignacion'")
            loConsulta.AppendLine("                            THEN Prestamos.Mon_Int")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        END)                        AS Prestamo_Mon_Int,")
            loConsulta.AppendLine("                    SUM(CASE ")
            loConsulta.AppendLine("                            WHEN Renglones_Recibos.Tipo = 'Deduccion'")
            loConsulta.AppendLine("                            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        END)                        AS Prestamo_Monto_Cuota")
            loConsulta.AppendLine("            FROM    Renglones_Recibos")
            loConsulta.AppendLine("                JOIN Prestamos ")
            loConsulta.AppendLine("                    ON  Prestamos.Documento = Renglones_Recibos.Doc_Ori")
            loConsulta.AppendLine("                    AND Prestamos.Cod_Tip BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("                LEFT JOIN Detalles_Prestamos ")
            loConsulta.AppendLine("                    ON  Detalles_Prestamos.Documento = Renglones_Recibos.Doc_Ori")
            loConsulta.AppendLine("                    AND Detalles_Prestamos.Renglon = Renglones_Recibos.Ren_Ori")
            loConsulta.AppendLine("                    AND Detalles_Prestamos.Pagado = 1 ")
            loConsulta.AppendLine("            WHERE   Renglones_Recibos.Tip_Ori = 'Prestamos'")
            loConsulta.AppendLine("            GROUP BY Renglones_Recibos.Documento, ")
            loConsulta.AppendLine("                    Prestamos.Documento, ")
            loConsulta.AppendLine("                    Prestamos.Motivo,")
            loConsulta.AppendLine("                    Prestamos.Fec_Asi, ")
            loConsulta.AppendLine("                    Prestamos.Cuotas, ")
            loConsulta.AppendLine("                    Renglones_Recibos.Ren_Ori, ")
            loConsulta.AppendLine("                    Detalles_Prestamos.Fec_Ven")
            loConsulta.AppendLine("	        ) Prestamo")
            loConsulta.AppendLine("        ON  Prestamo.Prestamo_Recibo = Recibos.Documento")
            loConsulta.AppendLine("    LEFT JOIN Motivos ")
            loConsulta.AppendLine("		ON	Motivos.Cod_Mot = Prestamo.Prestamo_Motivo")
            loConsulta.AppendLine("WHERE    Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("     AND Recibos.Status IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("     AND Recibos.Cod_Tra BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Con BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Car BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Suc BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("     AND Recibos.Cod_Rev BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loConsulta.AppendLine("ORDER BY	Recibos.Cod_Tra ASC, " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMovimientos_Prestamos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMovimientos_Prestamos.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 13/08/14: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
