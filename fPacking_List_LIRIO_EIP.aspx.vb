'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPacking_List_LIRIO_EIP"
'-------------------------------------------------------------------------------------------'
Partial Class fPacking_List_LIRIO_EIP
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT    Facturas.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")

            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Ent,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Ent,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Ent,1, 200) ELSE SUBSTRING(Facturas.Dir_Ent,1, 200) END) END) AS  Dir_Ent, ")

            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Facturas.Nom_Cli          As Nom_Gen, ")
            loConsulta.AppendLine("           Facturas.Rif              As Rif_Gen, ")
            loConsulta.AppendLine("           Facturas.Nit              As Nit_Gen, ")
            loConsulta.AppendLine("           Facturas.Dir_Fis          As Dir_Gen, ")
            loConsulta.AppendLine("           Facturas.Telefonos        As Tel_Gen, ")
            loConsulta.AppendLine("           Facturas.Documento,        ")
            loConsulta.AppendLine("           Facturas.Por_Des1         AS Por_Des1_Enc, ")
            loConsulta.AppendLine("           Facturas.Mon_Des1         AS Mon_Des1_Enc, ")
            loConsulta.AppendLine("           Facturas.Por_Rec1         AS Por_Rec1_Enc, ")
            loConsulta.AppendLine("           Facturas.Mon_Rec1         AS Mon_Rec1_Enc, ")
            loConsulta.AppendLine("           Renglones_Facturas.Cod_Uni, ")
            loConsulta.AppendLine("           Facturas.Fec_Ini, ")
            loConsulta.AppendLine("           Facturas.Cod_Mon, ")
            loConsulta.AppendLine("           Facturas.Fec_Fin, ")
            loConsulta.AppendLine("           Facturas.Mon_Bru, ")
            loConsulta.AppendLine("           Facturas.Por_Imp1, ")
            loConsulta.AppendLine("           Facturas.Dis_Imp, ")
            loConsulta.AppendLine("           Facturas.Mon_Imp1, ")
            loConsulta.AppendLine("           Facturas.Mon_Net, ")
            loConsulta.AppendLine("           Facturas.Cod_For, ")
            loConsulta.AppendLine("           Facturas.Comentario, ")
            loConsulta.AppendLine("           Facturas.Notas, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Facturas.For_Env, ")
            loConsulta.AppendLine("           Facturas.Tip_Env, ")
            loConsulta.AppendLine("           Facturas.Not_Des, ")
            loConsulta.AppendLine("           Facturas.Cod_Tra, ")
            loConsulta.AppendLine("           Transportes.Nom_Tra, ")
            loConsulta.AppendLine("           Facturas.Cod_Ven, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
            loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loConsulta.AppendLine("			ELSE Renglones_Facturas.Notas")
            loConsulta.AppendLine("		END														AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_Facturas.Renglon, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_Facturas.Renglon>26)")
            loConsulta.AppendLine("               THEN CHAR((Renglones_Facturas.Renglon-1) / 26 + 64)")
            loConsulta.AppendLine("               ELSE ''")
            loConsulta.AppendLine("           END) + CHAR(((Renglones_Facturas.Renglon-1) % 26 ) + 65)       AS Letra,")
            loConsulta.AppendLine("           Renglones_Facturas.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Facturas.Por_Des      As Por_Des1, ")
            loConsulta.AppendLine("           Renglones_Facturas.Precio1      As Precio1, ")
            loConsulta.AppendLine("           Renglones_Facturas.Comentario   As Comentario_Renglon, ")
            loConsulta.AppendLine("           Renglones_Facturas.Mon_Net      As Neto, ")
            loConsulta.AppendLine("           Renglones_Facturas.Cod_Imp      As Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Facturas.Por_Imp1     As Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Facturas.Mon_Imp1     As Impuesto ")
            loConsulta.AppendLine("FROM       Facturas ")
            loConsulta.AppendLine("    JOIN   Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Formas_Pagos.Cod_For = Facturas.Cod_For")
            loConsulta.AppendLine("    JOIN   Transportes ON Transportes.Cod_Tra = Facturas.Cod_Tra")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loConsulta.AppendLine("    JOIN   Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPacking_List_LIRIO_EIP", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPacking_List_LIRIO_EIP.ReportSource = loObjetoReporte

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
' RJG: 17/07/14: Código Inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 16/09/14: Se agregó el campo Notas con un recuadro. Se cambió "IP Doral" por         '
'                "Everything IP" en la nora al pié.                                         '
'-------------------------------------------------------------------------------------------'
' RJG: 13/12/14: Se ajustó el nombre del cliente para que ocupe 2 líneas si no cabe.        '
'-------------------------------------------------------------------------------------------'
