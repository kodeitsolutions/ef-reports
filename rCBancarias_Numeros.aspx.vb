Imports System.Data
Partial Class rCBancarias_Numeros

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Bancarias.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Status, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Sal_Act, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Sal_Con, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           (Case When Cuentas_Bancarias.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Bancarias, Bancos ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Bancarias.Cod_Cue     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("           And Cuentas_Bancarias.Status  IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("  ORDER BY      Cuentas_Bancarias." & lcOrdenamiento)
		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCBancarias_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCBancarias_Numeros.ReportSource = loObjetoReporte

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
' JJD: 02/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Ajuste de la Vista de Diseño
'-------------------------------------------------------------------------------------------'
