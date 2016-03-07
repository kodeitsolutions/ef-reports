'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_Giros"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_Giros
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
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Pagar.Documento, ")
			loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cuentas_Pagar.Fec_Ini, 103)	AS  Fec_Ini, ")
			loComandoSeleccionar.AppendLine("           YEAR(Cuentas_Pagar.Fec_Ini)                    AS  Anno, ")
			loComandoSeleccionar.AppendLine("           MONTH(Cuentas_Pagar.Fec_Ini)					AS  Mes, ")
			loComandoSeleccionar.AppendLine("           DAY(Cuentas_Pagar.Fec_Ini)						AS  Dia, ")
			loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Net                          AS  Mon_Net, ")
			loComandoSeleccionar.AppendLine("           SUBSTRING(Cuentas_Pagar.Comentario,1,10)       AS  Num_Gir, ")
			loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Comentario                       AS  Comentario, ")
			loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
			loComandoSeleccionar.AppendLine("           CAST('' AS CHAR(400))							AS  Mon_Let ")
			loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar, ")
			loComandoSeleccionar.AppendLine("           Proveedores ")
			loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Cod_Pro			=   Proveedores.Cod_Pro ")
			loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Tip		=   'GIRO' ")
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Documento	Between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Fec_Ini		Between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Pro		Between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Ven		Between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Status		IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Tra		Between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Mon		Between " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Rev		Between " & lcParametro7Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

			Dim lnMontoNumero As Decimal

			For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

				lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
				loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

			Next loFilas

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_Giros", laDatosReporte)

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrCPagar_Giros.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 25/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
