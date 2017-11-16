'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_fCotizaciones_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_fCotizaciones_Clientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	 Cotizaciones.Cod_Cli,")
            loComandoSeleccionar.AppendLine("		 Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("        Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("        Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("        Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Status,  ")
            loComandoSeleccionar.AppendLine("        CASE WHEN Cotizaciones.Control = ''")
            loComandoSeleccionar.AppendLine("             THEN '-'")
            loComandoSeleccionar.AppendLine("             ELSE Cotizaciones.Control")
            loComandoSeleccionar.AppendLine("        END                                AS Requisicion,  ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("        Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("        Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Comentario  AS Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Notas, ")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("        CASE WHEN Renglones_Cotizaciones.Cod_Uni = Renglones_Cotizaciones.Cod_Uni2 ")
            loComandoSeleccionar.AppendLine("             THEN Renglones_Cotizaciones.Cod_Uni")
            loComandoSeleccionar.AppendLine("             ELSE Renglones_Cotizaciones.Cod_Uni2")
            loComandoSeleccionar.AppendLine("        END                                AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Precio1,")
            loComandoSeleccionar.AppendLine("        Renglones_Cotizaciones.Mon_Net     AS  Neto ")
            loComandoSeleccionar.AppendLine("FROM Cotizaciones ")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Cotizaciones ON Cotizaciones.Documento = Renglones_Cotizaciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Cotizaciones.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("KDE_fCotizaciones_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvKDE_fCotizaciones_Clientes.ReportSource = loObjetoReporte

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
