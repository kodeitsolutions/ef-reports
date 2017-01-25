'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_fNota_Entrega"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_fNota_Entrega

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	 Entregas.Cod_Cli,")
            loComandoSeleccionar.AppendLine("		 Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("        Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("        Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("        Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("        Entregas.Documento, ")
            loComandoSeleccionar.AppendLine("        Entregas.Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("        Entregas.Status,  ")
            loComandoSeleccionar.AppendLine("        Entregas.Comentario, ")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("        Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Renglon, ")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Notas, ")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Doc_Ori     AS Cotizacion,")
            loComandoSeleccionar.AppendLine("        Entregas.Ord_Com               AS Orden_Compra,")
            loComandoSeleccionar.AppendLine("        Renglones_Entregas.Mon_Net     AS  Neto ")
            loComandoSeleccionar.AppendLine("FROM Entregas ")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Entregas ON Entregas.Documento = Renglones_Entregas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Entregas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("KDE_fNota_Entrega", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvKDE_fNota_Entrega.ReportSource = loObjetoReporte

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
