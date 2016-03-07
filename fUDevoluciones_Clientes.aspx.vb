'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUDevoluciones_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class fUDevoluciones_Clientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT TOP 40")
            loComandoSeleccionar.AppendLine(" 		Clientes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Clientes.Documento,")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Clientes.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DClientes.Cod_Art,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DClientes.Cod_Uni,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DClientes.Cod_Alm,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DClientes.Can_Art1,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DClientes.Precio1")
            loComandoSeleccionar.AppendLine(" FROM Clientes")
            loComandoSeleccionar.AppendLine(" JOIN Devoluciones_Clientes ON Devoluciones_Clientes.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_DClientes ON Devoluciones_Clientes.Documento = Renglones_DClientes.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_DClientes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Clientes.Status IN ('Confirmado','Afectado','Procesado')")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Devoluciones_Clientes.Fec_Ini DESC")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUDevoluciones_Clientes", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfUDevoluciones_Clientes.ReportSource = loObjetoReporte

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
' Douglas Cortez: 20/05/2010 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 23/05/2011: mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
