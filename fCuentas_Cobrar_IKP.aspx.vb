'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCuentas_Cobrar_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fCuentas_Cobrar_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Cuentas_Cobrar.Cod_Cli, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Cuentas_Cobrar.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cuentas_Cobrar.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Telefonos = '') THEN Clientes.Telefonos ELSE Cuentas_Cobrar.Telefonos END) END) AS  Telefonos, ")
            
            
            loConsulta.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Nom_Cli                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Documento, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Factura, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Control, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Mon_Bru, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Por_Imp1, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Tip_Ori, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Por_Des, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Mon_Des, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Por_Rec, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Mon_Rec, ")
            loConsulta.AppendLine("           (Cuentas_Cobrar.Mon_Otr1 + Cuentas_Cobrar.Mon_Otr2 + Cuentas_Cobrar.Mon_Otr3)   AS  Mon_Otr, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Mon_Net, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente  ")
            loConsulta.AppendLine("FROM       Cuentas_Cobrar  ")
            loConsulta.AppendLine("  JOIN     Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loConsulta.AppendLine("  JOIN     Formas_Pagos ON Cuentas_Cobrar.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("  JOIN     Vendedores ON Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = Cuentas_Cobrar.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Cobrar_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCuentas_Cobrar_IKP.ReportSource = loObjetoReporte

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
' RJG: 01/08/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
