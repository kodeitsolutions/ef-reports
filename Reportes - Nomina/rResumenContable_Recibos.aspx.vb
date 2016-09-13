'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumenContable_Recibos"
'-------------------------------------------------------------------------------------------'
Partial Class rResumenContable_Recibos
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

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	    Renglones_Recibos.Cod_Con					AS Cod_Con,			")
            loComandoSeleccionar.AppendLine("			Renglones_Recibos.Nom_con					AS Nom_con,			")
            loComandoSeleccionar.AppendLine("			Renglones_Recibos.tipo						AS Tipo,			")
            loComandoSeleccionar.AppendLine("			SUM(CASE Renglones_Recibos.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'Asignacion' THEN Renglones_Recibos.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("			END)													AS Mon_Asi,")
            loComandoSeleccionar.AppendLine("			SUM(CASE Renglones_Recibos.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'Deduccion' THEN -Renglones_Recibos.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("			END)													AS Mon_Ded,")
            loComandoSeleccionar.AppendLine("			SUM(CASE Renglones_Recibos.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'Retencion' THEN -Renglones_Recibos.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("			END)													AS Mon_Ret")
            loComandoSeleccionar.AppendLine("FROM		Recibos  ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recibos   ")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Recibos.documento = Recibos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Trabajadores   ")
            loComandoSeleccionar.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loComandoSeleccionar.AppendLine("WHERE		Renglones_Recibos.tipo <> 'Otro' ")
            loComandoSeleccionar.AppendLine("		AND	Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY    Renglones_Recibos.Cod_Con,")
            loComandoSeleccionar.AppendLine("            Renglones_Recibos.Nom_con,	")
            loComandoSeleccionar.AppendLine("            Renglones_Recibos.tipo		")
            loComandoSeleccionar.AppendLine("ORDER BY	(CASE Renglones_Recibos.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'Asignacion' THEN 0")
            loComandoSeleccionar.AppendLine("				WHEN 'Deduccion' THEN 1")
            loComandoSeleccionar.AppendLine("				WHEN 'Retencion' THEN 2")
            loComandoSeleccionar.AppendLine("				ELSE 3")
            loComandoSeleccionar.AppendLine("			END) ASC, ")
            loComandoSeleccionar.AppendLine("			Renglones_Recibos.Cod_Con")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumenContable_Recibos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumenContable_Recibos.ReportSource = loObjetoReporte

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
' RJG: 23/09/14: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
