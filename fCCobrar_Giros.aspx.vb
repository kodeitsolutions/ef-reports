Imports System.Data
Partial Class fCCobrar_Giros
	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Ini, 103)	AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           YEAR(Cuentas_Cobrar.Fec_Ini)                    AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Cuentas_Cobrar.Fec_Ini)					AS  Mes, ")
            loComandoSeleccionar.AppendLine("           DAY(Cuentas_Cobrar.Fec_Ini)						AS  Dia, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Net                          AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Cuentas_Cobrar.Comentario,1,10)       AS  Num_Gir, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Comentario                       AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine("           CAST('' AS CHAR(400))							AS  Mon_Let ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

			Dim lnMontoNumero As Decimal
			For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

				lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
				loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

			Next loFilas

			loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCCobrar_Giros", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvfCCobrar_Giros.ReportSource = loObjetoReporte

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
