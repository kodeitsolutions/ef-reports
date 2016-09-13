'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVacacionesRecibos_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rVacacionesRecibos_Fechas
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT  Recibos.Documento                            AS Documento,	")
            loConsulta.AppendLine("        Recibos.fecha                                AS Fecha,		")
            loConsulta.AppendLine("        Recibos.Status                               AS Estatus,		")
            loConsulta.AppendLine("        Recibos.Cod_Tra                              AS Cod_Tra,		")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra                         AS Nom_Tra,		")
            loConsulta.AppendLine("        COALESCE(Vacaciones.Fec_Ini,Recibos.Fec_Ini)	AS Fec_Ini,		")
            loConsulta.AppendLine("        COALESCE(Vacaciones.Fec_Fin,Recibos.Fec_Fin)	AS Fec_Fin,	    ")
            loConsulta.AppendLine("        COALESCE(Vacaciones.Dias,0)                  AS Dias,	    ")
            loConsulta.AppendLine("        Recibos.Mon_Asi                              AS Mon_Asi,		")
            loConsulta.AppendLine("        Recibos.Mon_Ded                              AS Mon_Ded,		")
            loConsulta.AppendLine("        Recibos.Mon_Ret                              AS Mon_Ret,		")
            loConsulta.AppendLine("        Recibos.Mon_Net                              AS Mon_Net,		")
            loConsulta.AppendLine("        COALESCE(Vacacion.Vacacion_Documento,'')     AS Vacacion_Documento,	")
            loConsulta.AppendLine("        COALESCE(Vacaciones.Motivo, '')              AS Cod_Mot,		")
            loConsulta.AppendLine("        COALESCE(Motivos.Nom_Mot, '')                AS Nom_Mot,		")
            loConsulta.AppendLine("        Recibos.Comentario                           AS Comentario	")
            loConsulta.AppendLine("FROM		Recibos ")
            loConsulta.AppendLine("    JOIN	Trabajadores ")
            loConsulta.AppendLine("        ON Trabajadores.cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN(  SELECT  Renglones_Recibos.Documento AS Vacacion_recibo,")
            loConsulta.AppendLine("                        Renglones_Recibos.Doc_Ori AS Vacacion_Documento")
            loConsulta.AppendLine("                FROM    Renglones_Recibos")
            loConsulta.AppendLine("	               WHERE   Renglones_Recibos.Tip_Ori = 'Vacaciones'")
            loConsulta.AppendLine("	               GROUP BY Renglones_Recibos.Documento, Renglones_Recibos.Doc_Ori")
            loConsulta.AppendLine("	        ) Vacacion")
            loConsulta.AppendLine("        ON  Vacacion.Vacacion_recibo = Recibos.Documento")
            loConsulta.AppendLine("    LEFT JOIN Vacaciones ")
            loConsulta.AppendLine("        ON Vacaciones.Documento = Vacacion.Vacacion_Documento")
            loConsulta.AppendLine("    LEFT JOIN Motivos ")
            loConsulta.AppendLine("		ON	Motivos.Cod_Mot = Vacaciones.Motivo")
            loConsulta.AppendLine("WHERE    Recibos.Cod_Con = '91'")
            loConsulta.AppendLine("	    AND Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("     AND Recibos.Status IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("     AND Recibos.Cod_Tra BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Con BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Car BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("     AND Trabajadores.Cod_Suc BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("     AND Recibos.Cod_Rev BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("ORDER BY	Recibos.Fecha ASC, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVacacionesRecibos_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrVacacionesRecibos_Fechas.ReportSource = loObjetoReporte

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
