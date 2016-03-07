﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPedidos_LIRIO"
'-------------------------------------------------------------------------------------------'
Partial Class fPedidos_LIRIO
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT    Pedidos.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Pedidos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Pedidos.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Pedidos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Proveedores.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Pedidos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")

            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Pedidos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Ent,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE SUBSTRING(Pedidos.Dir_Ent,1, 200) END) END) AS  Dir_Ent, ")

            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Pedidos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Proveedores.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Pedidos.Nom_Pro          As Nom_Gen, ")
            loConsulta.AppendLine("           Pedidos.Rif              As Rif_Gen, ")
            loConsulta.AppendLine("           Pedidos.Nit              As Nit_Gen, ")
            loConsulta.AppendLine("           Pedidos.Dir_Fis          As Dir_Gen, ")
            loConsulta.AppendLine("           Pedidos.Telefonos        As Tel_Gen, ")
            loConsulta.AppendLine("           Pedidos.Documento,        ")
            loConsulta.AppendLine("           Pedidos.Por_Des1         AS Por_Des1_Enc, ")
            loConsulta.AppendLine("           Pedidos.Mon_Des1         AS Mon_Des1_Enc, ")
            loConsulta.AppendLine("           Pedidos.Por_Rec1         AS Por_Rec1_Enc, ")
            loConsulta.AppendLine("           Pedidos.Mon_Rec1         AS Mon_Rec1_Enc, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loConsulta.AppendLine("           Pedidos.Fec_Ini, ")
            loConsulta.AppendLine("           Pedidos.Cod_Mon, ")
            loConsulta.AppendLine("           Pedidos.Fec_Fin, ")
            loConsulta.AppendLine("           Pedidos.Mon_Bru, ")
            loConsulta.AppendLine("           Pedidos.Por_Imp1, ")
            loConsulta.AppendLine("           Pedidos.Dis_Imp, ")
            loConsulta.AppendLine("           Pedidos.Mon_Imp1, ")
            loConsulta.AppendLine("           Pedidos.Mon_Net, ")
            loConsulta.AppendLine("           Pedidos.Cod_For, ")
            loConsulta.AppendLine("           Pedidos.Comentario, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Pedidos.Cod_Ven, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
            loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loConsulta.AppendLine("			ELSE Renglones_Pedidos.Notas")
            loConsulta.AppendLine("		END														AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_Pedidos.Renglon, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Por_Des      As Por_Des1, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Precio1      As Precio1, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Comentario   As Comentario_Renglon, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Mon_Net      As Neto, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Imp      As Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Por_Imp1     As Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Mon_Imp1     As Impuesto ")
            loConsulta.AppendLine("FROM       Pedidos ")
            loConsulta.AppendLine("    JOIN   Renglones_Pedidos ON Renglones_Pedidos.Documento = Pedidos.Documento")
            loConsulta.AppendLine("    JOIN   Proveedores ON Proveedores.Cod_Pro = Pedidos.Cod_Pro")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Formas_Pagos.Cod_For = Pedidos.Cod_For")
            loConsulta.AppendLine("    JOIN   Transportes ON Transportes.Cod_Tra = Pedidos.Cod_Tra")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Pedidos.Cod_Art")
            loConsulta.AppendLine("    JOIN   Vendedores ON Vendedores.Cod_Ven = Pedidos.Cod_Ven")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPedidos_LIRIO", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPedidos_LIRIO.ReportSource = loObjetoReporte

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
' RJG: 02/07/14: Código Inicial.                                                            '
'-------------------------------------------------------------------------------------------'
