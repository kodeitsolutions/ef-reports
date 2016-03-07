'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fProveedoresUltimas_20Facturas"
'-------------------------------------------------------------------------------------------'
Partial Class fProveedoresUltimas_20Facturas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT TOP 20")
            loComandoSeleccionar.AppendLine(" 		Compras.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		Compras.Documento,")
            loComandoSeleccionar.AppendLine(" 		Compras.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 		Compras.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 		Compras.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		Compras.Cod_For,")
            loComandoSeleccionar.AppendLine(" 		Compras.Cod_Tra,")
            loComandoSeleccionar.AppendLine(" 		Compras.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 		Compras.Mon_Sal,")
            loComandoSeleccionar.AppendLine(" 		Compras.Comentario,")
            loComandoSeleccionar.AppendLine(" 		Compras.Status")
            loComandoSeleccionar.AppendLine(" FROM Proveedores")
            loComandoSeleccionar.AppendLine(" JOIN Compras ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine("       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProveedoresUltimas_20Facturas", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfProveedoresUltimas_20Facturas.ReportSource = loObjetoReporte

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
