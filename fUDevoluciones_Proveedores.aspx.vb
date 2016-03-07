'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUDevoluciones_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fUDevoluciones_Proveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT TOP 40")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Proveedores.Documento,")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Proveedores.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Cod_Art,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Cod_Uni,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Cod_Alm,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Can_Art1,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Precio1")
            loComandoSeleccionar.AppendLine(" FROM Proveedores")
            loComandoSeleccionar.AppendLine(" JOIN Devoluciones_Proveedores ON Devoluciones_Proveedores.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_DProveedores ON Devoluciones_Proveedores.Documento = Renglones_DProveedores.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_DProveedores.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE")
            loComandoSeleccionar.AppendLine(" 		Devoluciones_Proveedores.Status IN ('Confirmado','Afectado','Procesado')")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Devoluciones_Proveedores.Fec_Ini DESC")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUDevoluciones_Proveedores", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfUDevoluciones_Proveedores.ReportSource = loObjetoReporte

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
' MAT: 23/05/2011: Ajuste del filtro Status y mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
