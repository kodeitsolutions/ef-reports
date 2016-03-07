'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fProformaInvoice_LIRIO"
'-------------------------------------------------------------------------------------------'
Partial Class fProformaInvoice_LIRIO
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT    Cotizaciones.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")

            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Ent,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Ent,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Ent,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Ent,1, 200) END) END) AS  Dir_Ent, ")

            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cotizaciones.Nom_Cli          As Nom_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Rif              As Rif_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Nit              As Nit_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Dir_Fis          As Dir_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Telefonos        As Tel_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Documento,        ")
            loConsulta.AppendLine("           Cotizaciones.Por_Des1         AS Por_Des1_Enc, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Des1         AS Mon_Des1_Enc, ")
            loConsulta.AppendLine("           Cotizaciones.Por_Rec1         AS Por_Rec1_Enc, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Rec1         AS Mon_Rec1_Enc, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            loConsulta.AppendLine("           Cotizaciones.Fec_Ini, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_Mon, ")
            loConsulta.AppendLine("           Cotizaciones.Fec_Fin, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Bru, ")
            loConsulta.AppendLine("           Cotizaciones.Por_Imp1, ")
            loConsulta.AppendLine("           Cotizaciones.Dis_Imp, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Net, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_For, ")
            loConsulta.AppendLine("           Cotizaciones.Comentario, ")
            loConsulta.AppendLine("           Cotizaciones.Notas, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Cotizaciones.For_Env, ")
            loConsulta.AppendLine("           Cotizaciones.Tip_Env, ")
            loConsulta.AppendLine("           Cotizaciones.Not_Des, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_Tra, ")
            loConsulta.AppendLine("           Transportes.Nom_Tra, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_Ven, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
            loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loConsulta.AppendLine("			ELSE Renglones_Cotizaciones.Notas")
            loConsulta.AppendLine("		END														AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Renglon>26)")
            loConsulta.AppendLine("               THEN CHAR((Renglones_Cotizaciones.Renglon-1) / 26 + 64)")
            loConsulta.AppendLine("               ELSE ''")
            loConsulta.AppendLine("           END) + CHAR(((Renglones_Cotizaciones.Renglon-1) % 26 ) + 65)       AS Letra,")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Por_Des      As Por_Des1, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Precio1      As Precio1, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Comentario   As Comentario_Renglon, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Mon_Net      As Neto, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Imp      As Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Por_Imp1     As Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Mon_Imp1     As Impuesto ")
            loConsulta.AppendLine("FROM       Cotizaciones ")
            loConsulta.AppendLine("    JOIN   Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Clientes.Cod_Cli = Cotizaciones.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Formas_Pagos.Cod_For = Cotizaciones.Cod_For")
            loConsulta.AppendLine("    JOIN   Transportes ON Transportes.Cod_Tra = Cotizaciones.Cod_Tra")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art")
            loConsulta.AppendLine("    JOIN   Vendedores ON Vendedores.Cod_Ven = Cotizaciones.Cod_Ven")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProformaInvoice_LIRIO", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfProformaInvoice_LIRIO.ReportSource = loObjetoReporte

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
' RJG: 09/07/14: Código Inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 18/07/14: Se cambió para que se cargue desde Cotizaciones (Originalmente cargaba     '
'                desde Cotizaciones de Venta.                                                   '
'-------------------------------------------------------------------------------------------'
' RJG: 16/09/14: Se agregó el campo Notas con un recuadro.                                  '
'-------------------------------------------------------------------------------------------'
' RJG: 13/12/14: Se ajustó el nombre del cliente para que ocupe 2 líneas si no cabe.        '
'-------------------------------------------------------------------------------------------'
