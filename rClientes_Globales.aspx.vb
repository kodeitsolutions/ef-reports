Imports System.Data
Partial Class rClientes_Globales

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Rif, ")
            loComandoSeleccionar.AppendLine("           Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Telefonos, ")
            loComandoSeleccionar.AppendLine("           (Case Status When  'A' Then 'Activo' When 'S' then 'Suspendido' Else 'Inactivo' End) as Status ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cod_Cli      Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		    AND Status   IN (" & lcParametro1Desde & ")")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Cli ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_Globales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrClientes_Globales.ReportSource = loObjetoReporte

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
' GCR: 01/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  31/08/09: Metodo de ordenamiento, verificacionde registros, Boton imprimir
'-------------------------------------------------------------------------------------------'