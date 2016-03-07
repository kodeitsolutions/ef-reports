'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUltimas_20Facturas"
'-------------------------------------------------------------------------------------------'
Partial Class fUltimas_20Facturas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT TOP 20")
            loComandoSeleccionar.AppendLine(" 		Facturas.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Documento,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Cod_For,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Cod_Tra,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Mon_Sal,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Comentario,")
            loComandoSeleccionar.AppendLine(" 		Facturas.Status")
            loComandoSeleccionar.AppendLine(" FROM Clientes")
            loComandoSeleccionar.AppendLine(" JOIN Facturas ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine("       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Facturas.Fec_Ini DESC")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUltimas_20Facturas", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfUltimas_20Facturas.ReportSource = loObjetoReporte

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
' Douglas Cortez: 22/04/2010 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'
