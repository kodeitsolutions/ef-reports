'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRecibos_Renglones_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class rRecibos_Renglones_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcCampoSueldo As String

            lcCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "")))

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Trabajadores.Cod_Tra	        AS Cod_Tra,	")
            loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	        AS Nom_Tra,	")
            loComandoSeleccionar.AppendLine("			Recibos.fec_ini                 AS desde,	")
            loComandoSeleccionar.AppendLine("			Recibos.fec_fin	                AS hasta,	")
            loComandoSeleccionar.AppendLine("			Trabajadores.cedula             AS cedula,")
            loComandoSeleccionar.AppendLine("			COALESCE(Renglones_Campos_Nomina.val_num, 0) AS Sueldo,")
            loComandoSeleccionar.AppendLine("			Recibos.Documento		        AS Documento,")
            loComandoSeleccionar.AppendLine("			Recibos.fecha			        AS Fecha,	")
            loComandoSeleccionar.AppendLine("			Conceptos_Nomina.cod_con        AS cod_con,")
            loComandoSeleccionar.AppendLine("			Conceptos_Nomina.nom_con        AS nom_con,")
            loComandoSeleccionar.AppendLine("			Renglones_Recibos.val_car       AS val_car,")
            loComandoSeleccionar.AppendLine("			Renglones_Recibos.val_num       AS val_num,")
            loComandoSeleccionar.AppendLine("        (CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("            WHEN 'Asignacion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("            ELSE 0")
            loComandoSeleccionar.AppendLine("        END)                                AS Asignacion,")
            loComandoSeleccionar.AppendLine("        (CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("            WHEN 'Deduccion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("            WHEN 'Retencion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("            ELSE 0")
            loComandoSeleccionar.AppendLine("        END)                                AS Deduccion	")
            loComandoSeleccionar.AppendLine("FROM		Recibos ")
            loComandoSeleccionar.AppendLine("    JOIN    Renglones_Recibos")
            loComandoSeleccionar.AppendLine("        ON  Renglones_Recibos.documento = Recibos.Documento")
            loComandoSeleccionar.AppendLine("        AND Renglones_Recibos.cod_con <> 'A013' ")
            loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
            loComandoSeleccionar.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loComandoSeleccionar.AppendLine("    JOIN    Conceptos_Nomina")
            loComandoSeleccionar.AppendLine("        ON  Conceptos_Nomina.Cod_Con = Renglones_Recibos.Cod_Con")
            loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.tipo <> 'Otro'")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Renglones_Campos_Nomina ")
            loComandoSeleccionar.AppendLine("        ON  Renglones_Campos_Nomina.cod_tra = Trabajadores.cod_tra")
            loComandoSeleccionar.AppendLine("        AND Renglones_Campos_Nomina.cod_cam = " & lcCampoSueldo)
            loComandoSeleccionar.AppendLine("WHERE		Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRecibos_Renglones_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRecibos_Renglones_GPV.ReportSource = loObjetoReporte

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
' EAG: 7/10/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
