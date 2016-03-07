'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fOrdenes_Cheque02"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fOrdenes_Cheque02
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Pagos.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nom_Pro           As  Nombre_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Rif               As  Rif_Genenerico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nit               As  Nit_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Dir_Fis           As  Dir_Fis_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Telefonos         As  Telefonos_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento         As  Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini           As  Fec_ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Fin           As  Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru           As  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp           As  Mon_Imp1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net           As  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret           As  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo            As  Motivo, ")

            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Renglon			As  Renglon_Cheque, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Num_Doc			As  Referencia, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Fec_ini			As  Fec_Ini_Cheque, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Mon_Net			As  Mon_Net_Cheque, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Cod_Cue			As  Cod_Cue_Che, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue		As  Nom_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue		As  Num_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban					As  Nom_Ban_Che, ")

            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con        As  Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon        As  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb        As  Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab        As  Mon_Hab, ")
            loComandoSeleccionar.AppendLine("            CASE ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net * -1 ")
            loComandoSeleccionar.AppendLine("            	ELSE ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net ")
            loComandoSeleccionar.AppendLine("            END       As  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("            CASE ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1 ")
            loComandoSeleccionar.AppendLine("            	ELSE ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("            END       As  Mon_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Imp        As  Cod_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Comentario     As  Comentario_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1       As  Mon_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))           As  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           Campos_Propiedades.Val_Car      As  Nom_Cob ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_oPagos           ON Ordenes_Pagos.Documento      =   Renglones_oPagos.Documento ")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores                ON Ordenes_Pagos.Cod_Pro        =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("   JOIN Conceptos                  ON Conceptos.Cod_Con            =   Renglones_oPagos.Cod_Con ")

            loComandoSeleccionar.AppendLine("   JOIN Detalles_Opagos            ON (Ordenes_Pagos.Documento     =   Detalles_Opagos.Documento ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Opagos.Tip_Ope =   'Cheque' ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Opagos.Renglon =   '1')")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Bancarias          ON (Cuentas_Bancarias.Cod_Cue   =   Detalles_Opagos.Cod_Cue) ")
            loComandoSeleccionar.AppendLine("   JOIN Bancos                     ON (Bancos.Cod_ban              =   cuentas_bancarias.Cod_Ban) ")

            loComandoSeleccionar.AppendLine("   LEFT JOIN Campos_Propiedades    ON Ordenes_Pagos.Documento      =   Campos_Propiedades.Cod_Reg ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Cod_Pro  = 'NOMCOBCHE1' ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Origen   = 'Ordenes_Pagos' ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)





            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fOrdenes_Cheque02", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fOrdenes_Cheque02.ReportSource = loObjetoReporte

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
' JFP: 28/11/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 03/05/10: Se Ajusto para tomar el debe sea (+) y el haber (-)
'-------------------------------------------------------------------------------------------'
' CMS: 30/06/10: Se Ajusto para tomar el Proveedor generico
'-------------------------------------------------------------------------------------------'
' JJD: 09/03/15: Se ajusto el formato para CGS
'-------------------------------------------------------------------------------------------'