Imports System.Data
Partial Class fComprobante_IVA

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Cuentas_Pagar.Documento                                                         AS Documento, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_Tip                                                           AS Cod_Tip, ")
            loConsulta.AppendLine("            (CASE WHEN Cuentas_Pagar.Cod_Tip = 'FACT' THEN Cuentas_Pagar.Factura ELSE '' END) AS Factura, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Control                                                           AS Control, ")
            loConsulta.AppendLine("            (CASE WHEN Cuentas_Pagar.Cod_Tip = 'N/CR' THEN Cuentas_Pagar.Factura ELSE '' END) AS Numero_NCR, ")
            loConsulta.AppendLine("            (CASE WHEN Cuentas_Pagar.Cod_Tip = 'N/DB' THEN Cuentas_Pagar.Factura ELSE '' END) AS Numero_NDB, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Num_Com                                                           AS Num_Com, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_Pro                                                           AS Cod_Pro, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_Ven                                                           AS Cod_Ven, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Status                                                            AS Status, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Fec_Ini                                                           AS Fec_Ini, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Fec_Fin                                                           AS Fec_Fin, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_Tra                                                           AS Cod_Tra, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_For                                                           AS Cod_For, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Cod_Mon                                                           AS Cod_Mon, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Tasa                                                              AS Tasa, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Bru                                                           AS Mon_Bru, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Exe                                                           AS Mon_Exe, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Por_Imp1                                                          AS Por_Imp, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Imp1                                                          AS Mon_Imp, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Net                                                           AS Mon_Net, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Sal                                                           AS Saldo, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Por_Rec                                                           AS Por_Rec, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Rec                                                           AS Mon_Rec, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Por_Des                                                           AS Por_Des, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Mon_Des                                                           AS Mon_Des, ")
            loConsulta.AppendLine("            (Cuentas_Pagar.Mon_Otr1 + Cuentas_Pagar.Mon_Otr2 + Cuentas_Pagar.Mon_Otr3)      AS Mon_Otr, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Comentario                                                        AS Comentario, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Tip_Ori                                                           AS Tip_Ori, ")
            loConsulta.AppendLine("            Cuentas_Pagar.Doc_Ori                                                           AS Doc_Ori, ")
            loConsulta.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("                (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("                (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("                (CASE WHEN (SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("                (CASE WHEN (Cuentas_Pagar.Telefonos = '') THEN Proveedores.Telefonos ELSE Cuentas_Pagar.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("            Proveedores.Fax                                                                 AS Fax, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                                                              AS Nom_Ven, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For                                                            AS Nom_For, ")
            loConsulta.AppendLine("            Transportes.Nom_Tra                                                             AS Nom_Tra, ")
            loConsulta.AppendLine("            CAST('01-REG' AS VARCHAR(10))                                                   AS Tipo_Transaccion, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Doc_Ori                                                  AS Ori_Doc, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Bas                                                  AS Ori_Obj, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Bas                                                  AS Ori_Bas, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Imp                                                  AS Ori_Imp, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Net                                                  AS Ori_Net, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Sus                                                  AS Ori_Sus, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Por_Ret                                                  AS Ori_Por, ")
            loConsulta.AppendLine("            Retenciones_Documentos.Mon_Ret                                                  AS Ori_Ret, ")
            loConsulta.AppendLine("            Origen_Pag.Factura                                                              AS Ori_Fac, ")
            loConsulta.AppendLine("            Origen_Pag.Control                                                              AS Ori_Con, ")
            loConsulta.AppendLine("            Origen_Pag.Cod_Tip                                                              AS Cod_Tip_Origen  ")
            loConsulta.AppendLine("FROM        Cuentas_Pagar")
            loConsulta.AppendLine("    JOIN    Proveedores    ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro ")
            loConsulta.AppendLine("    JOIN    Vendedores     ON Cuentas_Pagar.Cod_Ven = Vendedores.Cod_Ven ")
            loConsulta.AppendLine("    JOIN    Formas_Pagos   ON Cuentas_Pagar.Cod_For = Formas_Pagos.Cod_For ")
            loConsulta.AppendLine("    JOIN    Transportes    ON Cuentas_Pagar.Cod_Tra = Transportes.Cod_Tra ")
            loConsulta.AppendLine("    JOIN    Retenciones_Documentos ")
            loConsulta.AppendLine("        ON  Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento ")
            loConsulta.AppendLine("        AND Retenciones_Documentos.Cla_Des = Cuentas_Pagar.Cod_Tip ")
            loConsulta.AppendLine("        AND Retenciones_Documentos.Tip_Des = 'Cuentas_Pagar'")
            loConsulta.AppendLine("        AND Retenciones_Documentos.Clase = 'IMPUESTO'")
            loConsulta.AppendLine("        AND Retenciones_Documentos.Origen IN ('Pagos', 'Cuentas_Pagar')")
            loConsulta.AppendLine("    LEFT JOIN Cuentas_Pagar Origen_Pag")
            loConsulta.AppendLine("        ON  Origen_Pag.Documento = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("        AND Origen_Pag.Cod_Tip = Retenciones_Documentos.Cla_Ori")
            loConsulta.AppendLine("        AND Retenciones_Documentos.Tip_Ori = 'Cuentas_Pagar'")
            loConsulta.AppendLine("WHERE      Cuentas_Pagar.Status NOT IN ('Anulado') ")
            loConsulta.AppendLine("      AND  Cuentas_Pagar.Cod_Tip   = 'RETIVA' ")  
            loConsulta.AppendLine("      AND  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fComprobante_IVA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfComprobante_IVA.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' JJD: 24/07/09: Inclusion del porcentaje de impuesto
'-------------------------------------------------------------------------------------------'
' CMS: 16/03/10: Se Aplico el metodo de carga de imagen y									'
'				 el metodo de valiadación de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 15/04/10: Se ajusto para tomar el proveedor generico
'-------------------------------------------------------------------------------------------'
' JJD: 12/06/15: Adecuacion para Molven
'-------------------------------------------------------------------------------------------'
' RJG: 19/06/15: Continuación de la programación.                                           '
'-------------------------------------------------------------------------------------------'
