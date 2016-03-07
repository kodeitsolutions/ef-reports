'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCotizaciones_Clientes_StockMarket"
'-------------------------------------------------------------------------------------------'
Partial Class fCotizaciones_Clientes_StockMarket
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Cotizaciones.Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) AS  Rif, ")
            loConsulta.AppendLine("            Clientes.Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Cotizaciones.Dir_Fis = '') THEN Clientes.Dir_Fis ELSE Cotizaciones.Dir_Fis END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) AS  Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax, ")
            loConsulta.AppendLine("            Cotizaciones.Documento, ")
            loConsulta.AppendLine("            Cotizaciones.Fec_Ini, ")
            loConsulta.AppendLine("            Cotizaciones.Fec_Fin, ")
            loConsulta.AppendLine("            Cotizaciones.Cod_Mon, ")
            loConsulta.AppendLine("            Cotizaciones.Tasa, ")
            loConsulta.AppendLine("            Cotizaciones.Mon_Bru, ")
            loConsulta.AppendLine("            Cotizaciones.Por_Des1, ")
            loConsulta.AppendLine("            Cotizaciones.Por_Rec1, ")
            loConsulta.AppendLine("            Cotizaciones.Mon_Des1, ")
            loConsulta.AppendLine("            Cotizaciones.Mon_Rec1, ")
            loConsulta.AppendLine("            Cotizaciones.Mon_Imp1, ")
            loConsulta.AppendLine("            Cotizaciones.Mon_Net, ")
            loConsulta.AppendLine("            Cotizaciones.Cod_For, ")
            loConsulta.AppendLine("            Cotizaciones.Dis_Imp, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("            Cotizaciones.Cod_Ven, ")
            loConsulta.AppendLine("            Cotizaciones.Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Cod_Art, ")
            loConsulta.AppendLine("            (CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("                ELSE Renglones_Cotizaciones.Notas END) AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Renglon, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Comentario AS Comentario_renglon, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Can_Art1, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Precio1, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Por_Des, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Cotizaciones.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("            " & goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCorreo) & "         As  Correo_Empresa ")
            loConsulta.AppendLine("FROM        Cotizaciones  ")
            loConsulta.AppendLine("    JOIN    Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
            loConsulta.AppendLine("    JOIN    Clientes ON Clientes.Cod_Cli = Cotizaciones.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = Cotizaciones.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Cotizaciones.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos  ON Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art")
            loConsulta.AppendLine("WHERE       "  & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_Clientes_StockMarket", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCotizaciones_Clientes_StockMarket.ReportSource = loObjetoReporte

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
' RJG: 25/02/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
