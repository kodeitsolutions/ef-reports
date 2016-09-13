'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReposos"
'-------------------------------------------------------------------------------------------'
Partial Class rReposos
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

            loComandoSeleccionar.AppendLine("SELECT		Reposos.Documento	    AS Documento,	")
            loComandoSeleccionar.AppendLine("			Reposos.fecha		    AS Fecha,		")
            loComandoSeleccionar.AppendLine("			Reposos.[Status]		AS Estatus,		")
            loComandoSeleccionar.AppendLine("			Reposos.Cod_Tra		    AS Cod_Tra,		")
            loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	AS Nom_Tra,		")
            loComandoSeleccionar.AppendLine("			Reposos.Fec_Ini		    AS Fec_Ini,		")
            loComandoSeleccionar.AppendLine("			Reposos.Fec_Fin		    AS Fec_Fin,		")
            loComandoSeleccionar.AppendLine("			Reposos.Dias			AS Dias,		")
            loComandoSeleccionar.AppendLine("			Reposos.Cod_Rev		    AS Cod_Rev,		")
            loComandoSeleccionar.AppendLine("			Reposos.Motivo		    AS Motivo,		")
            loComandoSeleccionar.AppendLine("			Reposos.Comentario	    AS Comentario	")
            loComandoSeleccionar.AppendLine("FROM		Reposos ")
            loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
            loComandoSeleccionar.AppendLine("		ON	Trabajadores.cod_Tra = Reposos.Cod_Tra")
            loComandoSeleccionar.AppendLine("WHERE		Reposos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Reposos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Reposos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Reposos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Reposos.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReposos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrReposos.ReportSource = loObjetoReporte

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
' RJG: 16/02/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
