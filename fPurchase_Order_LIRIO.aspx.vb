'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPurchase_Order_LIRIO"
'-------------------------------------------------------------------------------------------'
Partial Class fPurchase_Order_LIRIO
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")

            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Ent,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Ent,1, 200) END) END) AS  Dir_Ent, ")

            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Ordenes_Compras.Nom_Pro          As Nom_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Rif              As Rif_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Nit              As Nit_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Dir_Fis          As Dir_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Telefonos        As Tel_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Documento,        ")
            loConsulta.AppendLine("           Ordenes_Compras.Por_Des1         AS Por_Des1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Des1         AS Mon_Des1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Por_Rec1         AS Por_Rec1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Rec1         AS Mon_Rec1_Enc, ")
            loConsulta.AppendLine("           Renglones_oCompras.Cod_Uni, ")
            loConsulta.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_Mon, ")
            loConsulta.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            loConsulta.AppendLine("           Ordenes_Compras.Por_Imp1, ")
            loConsulta.AppendLine("           Ordenes_Compras.Dis_Imp, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Net, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_For, ")
            loConsulta.AppendLine("           Ordenes_Compras.Comentario, ")
            loConsulta.AppendLine("           Ordenes_Compras.Notas, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Ordenes_Compras.For_Env, ")
            loConsulta.AppendLine("           Ordenes_Compras.Tip_Env, ")
            loConsulta.AppendLine("           Ordenes_Compras.Not_Des, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_Tra, ")
            loConsulta.AppendLine("           Transportes.Nom_Tra, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_oCompras.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
            loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loConsulta.AppendLine("			ELSE Renglones_oCompras.Notas")
            loConsulta.AppendLine("		END														AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_oCompras.Renglon, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_oCompras.Renglon>26)")
            loConsulta.AppendLine("               THEN CHAR((Renglones_oCompras.Renglon-1) / 26 + 64)")
            loConsulta.AppendLine("               ELSE ''")
            loConsulta.AppendLine("           END) + CHAR(((Renglones_oCompras.Renglon-1) % 26 ) + 65)       AS Letra,")
            loConsulta.AppendLine("           Renglones_oCompras.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_oCompras.Por_Des      As Por_Des1, ")
            loConsulta.AppendLine("           Renglones_oCompras.Precio1      As Precio1, ")
            loConsulta.AppendLine("           Renglones_oCompras.Comentario   As Comentario_Renglon, ")
            loConsulta.AppendLine("           Renglones_oCompras.Mon_Net      As Neto, ")
            loConsulta.AppendLine("           Renglones_oCompras.Cod_Imp      As Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_oCompras.Por_Imp1     As Por_Imp, ")
            loConsulta.AppendLine("           Renglones_oCompras.Mon_Imp1     As Impuesto ")
            loConsulta.AppendLine("FROM       Ordenes_Compras ")
            loConsulta.AppendLine("    JOIN   Renglones_oCompras ON Renglones_oCompras.Documento = Ordenes_Compras.Documento")
            loConsulta.AppendLine("    JOIN   Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Formas_Pagos.Cod_For = Ordenes_Compras.Cod_For")
            loConsulta.AppendLine("    JOIN   Transportes ON Transportes.Cod_Tra = Ordenes_Compras.Cod_Tra")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_oCompras.Cod_Art")
            loConsulta.AppendLine("    JOIN   Vendedores ON Vendedores.Cod_Ven = Ordenes_Compras.Cod_Ven")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPurchase_Order_LIRIO", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPurchase_Order_LIRIO.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 18/07/14: Código Inicial, a partir de fPedidos_LIRIO.                                '
'-------------------------------------------------------------------------------------------'
' RJG: 13/12/14: Se ajustó el nombre del cliente para que ocupe 2 líneas si no cabe.        '
'-------------------------------------------------------------------------------------------'
