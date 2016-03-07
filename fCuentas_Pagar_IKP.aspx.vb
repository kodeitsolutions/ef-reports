Imports System.Data
Partial Class fCuentas_Pagar_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Cuentas_Pagar.Documento, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_Tip, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Factura, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Control, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_Ven, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Status, ")
            loConsulta.AppendLine("           CONVERT(NCHAR(10), Cuentas_Pagar.Fec_Ini, 103)                                 AS  Fec_Ini, ")
            loConsulta.AppendLine("           CONVERT(NCHAR(10), Cuentas_Pagar.Fec_Fin, 103)                                 AS  Fec_Fin, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_Tra                                                          AS  Cod_Tra, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_For                                                          AS  Cod_For, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Cod_Mon                                                          AS  Cod_Mon, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Tasa                                                             AS  Tasa, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Bru                                                          AS  Mon_Bru, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Por_Imp1                                                         AS  Por_Imp, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Imp1                                                         AS  Mon_Imp, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Net                                                          AS  Mon_Net, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Sal                                                          AS  Saldo, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Por_Rec                                                          AS  Por_Rec, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Rec                                                          AS  Mon_Rec, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Por_Des                                                          AS  Por_Des, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Mon_Des                                                          AS  Mon_Des, ")
            loConsulta.AppendLine("           (Cuentas_Pagar.Mon_Otr1 + Cuentas_Pagar.Mon_Otr2 + Cuentas_Pagar.Mon_Otr3)     AS  Mon_Otr, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Comentario                                                       AS  Comentario, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Tip_Ori, ")
            loConsulta.AppendLine("           Cuentas_Pagar.Doc_Ori, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Pagar.Telefonos = '') THEN Proveedores.Telefonos ELSE Cuentas_Pagar.Telefonos END) END) AS  Telefonos, ")
            
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Transportes.Nom_Tra, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente  ")
            loConsulta.AppendLine("FROM       Cuentas_Pagar ")
            loConsulta.AppendLine("  JOIN     Proveedores ON Cuentas_Pagar.Cod_Pro      =   Proveedores.Cod_Pro")
            loConsulta.AppendLine("  JOIN     Vendedores ON Cuentas_Pagar.Cod_Ven  =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("  JOIN     Formas_Pagos ON Cuentas_Pagar.Cod_For  =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("  JOIN     Transportes ON Cuentas_Pagar.Cod_Tra  =   Transportes.Cod_Tra")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = Cuentas_Pagar.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Pagar_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Pagar_IKP.ReportSource = loObjetoReporte

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
' RJG: 02/08/13: Programacion inicial
'-------------------------------------------------------------------------------------------'
