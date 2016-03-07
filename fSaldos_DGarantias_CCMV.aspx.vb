'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSaldos_DGarantias_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class fSaldos_DGarantias_CCMV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    Clientes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			CASE")
            loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'A' THEN 'Activo'")
            loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'S' THEN 'Suspendido'")
            loComandoSeleccionar.AppendLine(" 			END AS Status,")
            loComandoSeleccionar.AppendLine(" 			Clientes.Rif,")
            loComandoSeleccionar.AppendLine(" 			Movimientos_Cuentas.Mon_Deb,")
            loComandoSeleccionar.AppendLine(" 			Movimientos_Cuentas.Mon_Hab")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporal")
            loComandoSeleccionar.AppendLine(" FROM      Clientes JOIN Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("           ON Movimientos_Cuentas.Cod_Reg =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Status      =   'Confirmado'")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cla_Doc =   'Cliente'")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            loComandoSeleccionar.AppendLine(" SELECT    Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Status,")
            loComandoSeleccionar.AppendLine(" 			Rif,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Deb)            AS  Mon_Deb,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Hab)            AS  Mon_Hab,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Deb - Mon_Hab)  AS  Saldo")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTemporal")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Status,")
            loComandoSeleccionar.AppendLine(" 			Rif")
            loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Status,")
            loComandoSeleccionar.AppendLine(" 			Rif")


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSaldos_DGarantias_CCMV", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfSaldos_DGarantias_CCMV.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 09/06/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
