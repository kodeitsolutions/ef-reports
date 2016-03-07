'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFactura_Venta_StockMarket"
'-------------------------------------------------------------------------------------------'
Partial Class fFactura_Venta_StockMarket
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Facturas.Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) AS  Rif, ")
            loConsulta.AppendLine("            Clientes.Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Dir_Fis = '') THEN Clientes.Dir_Fis ELSE Facturas.Dir_Fis END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) AS  Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax, ")
            loConsulta.AppendLine("            Facturas.Documento, ")
            loConsulta.AppendLine("            Facturas.Fec_Ini, ")
            loConsulta.AppendLine("            Facturas.Fec_Fin, ")
            loConsulta.AppendLine("            Facturas.Cod_Mon, ")
            loConsulta.AppendLine("            Facturas.Tasa, ")
            loConsulta.AppendLine("            Facturas.Mon_Bru, ")
            loConsulta.AppendLine("            Facturas.Por_Des1, ")
            loConsulta.AppendLine("            Facturas.Por_Rec1, ")
            loConsulta.AppendLine("            Facturas.Mon_Des1, ")
            loConsulta.AppendLine("            Facturas.Mon_Rec1, ")
            loConsulta.AppendLine("            Facturas.Mon_Imp1, ")
            loConsulta.AppendLine("            Facturas.Mon_Net, ")
            loConsulta.AppendLine("            Facturas.Cod_For, ")
            loConsulta.AppendLine("            Facturas.Dis_Imp, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("            Facturas.Cod_Ven, ")
            loConsulta.AppendLine("            Facturas.Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Art, ")
            loConsulta.AppendLine("            (CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("                ELSE Renglones_Facturas.Notas END) AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Facturas.Renglon, ")
            loConsulta.AppendLine("            Renglones_Facturas.Comentario AS Comentario_renglon, ")
            loConsulta.AppendLine("            Renglones_Facturas.Can_Art1, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_Facturas.Precio1, ")
            loConsulta.AppendLine("            Renglones_Facturas.Por_Des, ")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("            Renglones_Facturas.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("            " & goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCorreo) & "         As  Correo_Empresa ")
            loConsulta.AppendLine("FROM        Facturas  ")
            loConsulta.AppendLine("    JOIN    Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loConsulta.AppendLine("    JOIN    Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = Facturas.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos  ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFactura_Venta_StockMarket", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFactura_Venta_StockMarket.ReportSource = loObjetoReporte

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
