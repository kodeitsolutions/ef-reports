'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fEntrega_Albaran_StockMarket"
'-------------------------------------------------------------------------------------------'
Partial Class fEntrega_Albaran_StockMarket
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Entregas.Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Entregas.Nom_Cli END) AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Rif = '') THEN Clientes.Rif ELSE Entregas.Rif END) AS  Rif, ")
            loConsulta.AppendLine("            Clientes.Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Dir_Fis = '') THEN Clientes.Dir_Fis ELSE Entregas.Dir_Fis END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Telefonos = '') THEN Clientes.Telefonos ELSE Entregas.Telefonos END) AS  Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax, ")
            loConsulta.AppendLine("            Entregas.Documento, ")
            loConsulta.AppendLine("            Entregas.Fec_Ini, ")
            loConsulta.AppendLine("            Entregas.Fec_Fin, ")
            loConsulta.AppendLine("            Entregas.Cod_Mon, ")
            loConsulta.AppendLine("            Entregas.Tasa, ")
            loConsulta.AppendLine("            Entregas.Mon_Bru, ")
            loConsulta.AppendLine("            Entregas.Por_Des1, ")
            loConsulta.AppendLine("            Entregas.Por_Rec1, ")
            loConsulta.AppendLine("            Entregas.Mon_Des1, ")
            loConsulta.AppendLine("            Entregas.Mon_Rec1, ")
            loConsulta.AppendLine("            Entregas.Mon_Imp1, ")
            loConsulta.AppendLine("            Entregas.Mon_Net, ")
            loConsulta.AppendLine("            Entregas.Cod_For, ")
            loConsulta.AppendLine("            Entregas.Dis_Imp, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("            Entregas.Cod_Ven, ")
            loConsulta.AppendLine("            Entregas.Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Entregas.Cod_Art, ")
            loConsulta.AppendLine("            (CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("                ELSE Renglones_Entregas.Notas END) AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Entregas.Renglon, ")
            loConsulta.AppendLine("            Renglones_Entregas.Comentario AS Comentario_renglon, ")
            loConsulta.AppendLine("            Renglones_Entregas.Can_Art1, ")
            loConsulta.AppendLine("            Renglones_Entregas.Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_Entregas.Precio1, ")
            loConsulta.AppendLine("            Renglones_Entregas.Por_Des, ")
            loConsulta.AppendLine("            Renglones_Entregas.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("            Renglones_Entregas.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Entregas.Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Entregas.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("            " & goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCorreo) & "         As  Correo_Empresa ")
            loConsulta.AppendLine("FROM        Entregas  ")
            loConsulta.AppendLine("    JOIN    Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
            loConsulta.AppendLine("    JOIN    Clientes ON Clientes.Cod_Cli = Entregas.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = Entregas.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Entregas.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos  ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fEntrega_Albaran_StockMarket", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfEntrega_Albaran_StockMarket.ReportSource = loObjetoReporte

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
