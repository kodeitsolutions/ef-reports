'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUPrecios_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class fUPrecios_Compras
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT TOP 40")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		Compras.Documento,")
            loComandoSeleccionar.AppendLine(" 		Compras.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Cod_Art,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Cod_Uni,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Cod_Alm,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Can_Art1,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Precio1")
            loComandoSeleccionar.AppendLine(" FROM Proveedores")
            loComandoSeleccionar.AppendLine(" JOIN Compras ON Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Compras ON Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE")
            loComandoSeleccionar.AppendLine(" 		Compras.Status IN ('Afectado','Confirmado','Procesado')")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Compras.Fec_Ini DESC")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUPrecios_Compras", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfUPrecios_Compras.ReportSource = loObjetoReporte

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
' MAT: 23/05/2011: mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
