'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLiquidaciones"
'-------------------------------------------------------------------------------------------'
Partial Class rLiquidaciones
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Liquidaciones.Documento	    AS Documento,	")
            loComandoSeleccionar.AppendLine("			Liquidaciones.fecha		    AS Fecha,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.[Status]		AS Estatus,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Cod_Tra		AS Cod_Tra,		")
            loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	    AS Nom_Tra,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Fec_Ini		AS Fec_Ini,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Fec_Fin		AS Fec_Fin,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Cod_Rev		AS Cod_Rev,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Motivo		AS Motivo,		")
            loComandoSeleccionar.AppendLine("			Motivos.Nom_Mot	        	AS Nom_Mot,		")
            loComandoSeleccionar.AppendLine("			Liquidaciones.Comentario	AS Comentario	")
            loComandoSeleccionar.AppendLine("FROM		Liquidaciones ")
            loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
            loComandoSeleccionar.AppendLine("		ON	Trabajadores.cod_Tra = Liquidaciones.Cod_Tra")
            loComandoSeleccionar.AppendLine("	JOIN	Motivos ")
            loComandoSeleccionar.AppendLine("		ON	Motivos.Cod_Mot = Liquidaciones.Motivo")
            loComandoSeleccionar.AppendLine("WHERE		Liquidaciones.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Liquidaciones.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Liquidaciones.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Liquidaciones.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Liquidaciones.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLiquidaciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLiquidaciones.ReportSource = loObjetoReporte

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
' RJG: 10/04/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
