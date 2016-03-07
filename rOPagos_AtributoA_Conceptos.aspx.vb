'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_AtributoA_Conceptos"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_AtributoA_Conceptos
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
            
            Dim loConsulta As New StringBuilder()
            
			loConsulta.AppendLine("SELECT		CASE	WHEN (Conceptos.Atributo_A='')	")
			loConsulta.AppendLine("					THEN '(Sin Atributo A asignado)'")
			loConsulta.AppendLine("					ELSE Conceptos.Atributo_A		")
			loConsulta.AppendLine("			END AS Atributo_A,")
			loConsulta.AppendLine(" 			Conceptos.Cod_Con,")
			loConsulta.AppendLine(" 			Conceptos.Nom_con,")
			loConsulta.AppendLine(" 			SUM(Renglones_oPagos.Mon_deb)	AS Mon_Deb,")
			loConsulta.AppendLine(" 			SUM(Renglones_oPagos.Mon_hab)	AS Mon_Hab,")
			loConsulta.AppendLine(" 			SUM(Renglones_oPagos.Mon_imp1)	AS Mon_Imp1")
			loConsulta.AppendLine("FROM		Ordenes_Pagos ")
			loConsulta.AppendLine("	JOIN	Renglones_oPagos	ON Ordenes_Pagos.Documento = Renglones_oPagos.Documento")
			loConsulta.AppendLine("	JOIN	Conceptos			ON Renglones_oPagos.Cod_Con = Conceptos.Cod_con")
            loConsulta.AppendLine("		AND Conceptos.Atributo_A  BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("			AND " & lcParametro8Hasta)
			loConsulta.AppendLine("WHERE		Ordenes_Pagos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("				AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Fec_Ini   BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("				AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Cod_Pro   BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("				AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loConsulta.AppendLine("			AND Renglones_oPagos.Cod_Con  BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("				AND " & lcParametro4Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Cod_Mon  BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("				AND " & lcParametro5Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Cod_Rev  BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("				AND " & lcParametro6Hasta)
            loConsulta.AppendLine("			AND Ordenes_Pagos.Cod_Suc  BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("				AND " & lcParametro7Hasta)
			loConsulta.AppendLine("GROUP BY		Conceptos.Atributo_A,")
			loConsulta.AppendLine("				Conceptos.Cod_Con,")
			loConsulta.AppendLine("				Conceptos.Nom_Con")
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_AtributoA_Conceptos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_AtributoA_Conceptos.ReportSource = loObjetoReporte

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
' RJG: 23/03/15: Código inicial, a partir de OPagos_gConceptos.								'
'-------------------------------------------------------------------------------------------'
