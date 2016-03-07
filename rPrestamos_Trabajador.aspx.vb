'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrestamos_Trabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rPrestamos_Trabajador
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

            loComandoSeleccionar.AppendLine("SELECT		Prestamos.Documento	                AS Documento,	")
            loComandoSeleccionar.AppendLine("			Prestamos.fec_asi		            AS fec_asi,		")
            loComandoSeleccionar.AppendLine("			Prestamos.[Status]		            AS Estatus,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Cod_Tra		            AS Cod_Tra,		")
            loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	            AS Nom_Tra,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Mon_Bas		            AS Mon_Bas,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Por_Int		            AS Por_Int,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Mon_Int		            AS Mon_Int,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Mon_Net		            AS Mon_Net,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Mon_Pag		            AS Mon_Pag,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Mon_Sal		            AS Mon_Sal,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Cuotas		            AS Cuotas,		")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Prestamos.Cuotas>0)")
            loComandoSeleccionar.AppendLine("			    THEN Prestamos.Mon_Net/Prestamos.Cuotas")
            loComandoSeleccionar.AppendLine("			    ELSE Prestamos.Mon_Net")
            loComandoSeleccionar.AppendLine("			END)	                           AS Monto_Cuota,")
            loComandoSeleccionar.AppendLine("			Prestamos.Cod_Rev		            AS Cod_Rev,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Motivo		            AS Motivo,		")
            loComandoSeleccionar.AppendLine("			Prestamos.Comentario	            AS Comentario	")
            loComandoSeleccionar.AppendLine("FROM		Prestamos ")
            loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
            loComandoSeleccionar.AppendLine("		ON	Trabajadores.cod_Tra = Prestamos.Cod_Tra")
            loComandoSeleccionar.AppendLine("WHERE		Prestamos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Prestamos.fec_asi BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Prestamos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Prestamos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Prestamos.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPrestamos_Trabajador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPrestamos_Trabajador.ReportSource = loObjetoReporte

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
' RJG: 16/02/13: Codigo inicial.																'
'-------------------------------------------------------------------------------------------'
' RJG: 03/02/14: Se agregó la columna Pagado, y los totales de los montos.																'
'-------------------------------------------------------------------------------------------'
