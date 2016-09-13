'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDiario_General"
'-------------------------------------------------------------------------------------------'
Partial Class rDiario_General
	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
			
			loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Resumen, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Tipo, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Origen, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Integracion, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Status, ")
			loComandoSeleccionar.AppendLine("			Comprobantes.Notas, ")
			loComandoSeleccionar.AppendLine("			(CASE WHEN (Renglones_Comprobantes.Mon_Deb > 0) THEN 0 ELSE 1 END) AS Debe_Haber, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Renglon, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Fec_Ini						AS Fec_Ini_Renglon, ")
			loComandoSeleccionar.AppendLine("			DATEPART(DAY, Renglones_Comprobantes.Fec_Ini)		AS Dia_Ori, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Cen, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Gas, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Act, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Tip, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Cla, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Mon, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Tasa, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Mon_Deb, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Mon_Hab, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Comentario, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Referencia, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Tip_Ori, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Doc_Ori, ")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Reg ")
			loComandoSeleccionar.AppendLine("FROM		Comprobantes ")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Comprobantes ")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("WHERE		Comprobantes.Documento                  BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("           AND Renglones_Comprobantes.Fec_Ini		BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("           AND Renglones_Comprobantes.Cod_Mon      BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	Renglones_Comprobantes.Fec_Ini ASC,")
			loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Documento ASC")

 			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDiario_General", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrCDiario_Consecutivo.ReportSource = loObjetoReporte

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
' Fin del codigo.																			'
'-------------------------------------------------------------------------------------------'
' RJG:  21/10/11: Código inicial.															'
'-------------------------------------------------------------------------------------------'
