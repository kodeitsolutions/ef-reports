'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_gConceptos"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_gConceptos

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            
            Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		CASE	WHEN (Conceptos.Grupo='')	")
			loComandoSeleccionar.AppendLine("					THEN '(Sin grupo asignado)'	")
			loComandoSeleccionar.AppendLine("					ELSE Conceptos.Grupo		")
			loComandoSeleccionar.AppendLine("			END AS Grupo,")
			loComandoSeleccionar.AppendLine(" 			Conceptos.Cod_Con,")
			loComandoSeleccionar.AppendLine(" 			Conceptos.Nom_con,")
			loComandoSeleccionar.AppendLine(" 			SUM(Renglones_oPagos.Mon_deb)	AS Mon_Deb,")
			loComandoSeleccionar.AppendLine(" 			SUM(Renglones_oPagos.Mon_hab)	AS Mon_Hab,")
			loComandoSeleccionar.AppendLine(" 			SUM(Renglones_oPagos.Mon_imp1)	AS Mon_Imp1")
			loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos ")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_oPagos	ON Ordenes_Pagos.Documento = Renglones_oPagos.Documento")
			loComandoSeleccionar.AppendLine("	JOIN	Conceptos			ON Renglones_oPagos.Cod_Con = Conceptos.Cod_con")
            loComandoSeleccionar.AppendLine("		AND Conceptos.Grupo  BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Fec_Ini   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Pro   BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Renglones_oPagos.Cod_Con  BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Mon  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Rev  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Suc  BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY		Conceptos.Grupo,")
			loComandoSeleccionar.AppendLine("				Conceptos.Cod_Con,")
			loComandoSeleccionar.AppendLine("				Conceptos.Nom_Con")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_gConceptos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_gConceptos.ReportSource = loObjetoReporte

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
' RJG: 11/03/09: Código inicial																'
'-------------------------------------------------------------------------------------------'
