Imports System.Data
Partial Class rCuentas_Bancarias

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append(" SELECT    Cuentas_Bancarias.Cod_Cue, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Status, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Cod_Ban, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Sucursal, ")
            loComandoSeleccionar.Append("           Cuentas_Bancarias.Telefonos, ")
            loComandoSeleccionar.Append("           Bancos.Nom_Ban, ")
            loComandoSeleccionar.Append("           (Case When Cuentas_Bancarias.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Cuentas_Bancarias ")
            loComandoSeleccionar.Append(" FROM      Cuentas_Bancarias, Bancos ")
            loComandoSeleccionar.Append(" WHERE     Cuentas_Bancarias.Cod_Cue     Between " & lcParametro0Desde)
            loComandoSeleccionar.Append("           And " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           And Cuentas_Bancarias.Status  IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.Append("           And Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("  ORDER BY      Cuentas_Bancarias." & lcOrdenamiento)
    
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCuentas_Bancarias", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCuentas_Bancarias.ReportSource = loObjetoReporte

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
' MJP:  08/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP:  11/07/08: Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP:  14/07/08: Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD:  02/02/09: Se agrego la lista del Status
'-------------------------------------------------------------------------------------------'