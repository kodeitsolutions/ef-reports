Imports System.Data
'--------------------------------------------------------------------------------------------
Partial Class CSJ_fOrdenes_Cheque_BanCaribe
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit AS Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")

            loComandoSeleccionar.AppendLine("           convert(varchar, Ordenes_Pagos.Fec_Ini, 110) 			As  Fecha_Texto, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento         As  Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini           As  Fec_ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Fin           As  Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru           As  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp           As  Mon_Imp1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net           As  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret           As  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo            As  Motivo, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con        As  Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con, ")



            loComandoSeleccionar.AppendLine("           detalles_opagos.Renglon			As  Renglon_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_opagos.Num_Doc			As  Referencia, ")
            loComandoSeleccionar.AppendLine("           detalles_opagos.Fec_ini			As  Fec_Ini_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_opagos.Mon_Net			As  Mon_Net_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_opagos.Cod_Cue			As  Cod_Cue_Che, ")
            loComandoSeleccionar.AppendLine("           cuentas_bancarias.Nom_Cue		As  Nom_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           bancos.Nom_Ban					As  Nom_Ban_Che, ")

            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb        As  Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab        As  Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           CASE ")
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
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon     As  Renglon,")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Comentario     As  Comentario_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1       As  Mon_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))           As  Mon_Let ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_oPagos ON (Ordenes_Pagos.Documento =   Renglones_oPagos.Documento)")
            loComandoSeleccionar.AppendLine(" JOIN detalles_opagos ON (Ordenes_Pagos.Documento =   detalles_opagos.Documento AND detalles_opagos.tip_ope = 'Cheque' AND detalles_opagos.Renglon = '1')")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON (Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro) ")
            loComandoSeleccionar.AppendLine(" JOIN cuentas_bancarias ON (cuentas_bancarias.Cod_Cue   =   detalles_opagos.Cod_Cue) ")
            loComandoSeleccionar.AppendLine(" JOIN bancos ON (bancos.Cod_ban   =   cuentas_bancarias.Cod_Ban) ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Conceptos ON (Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con)")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")



            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CSJ_fOrdenes_Cheque_BanCaribe", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_fOrdenes_Cheque_BanCaribe.ReportSource = loObjetoReporte

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
' MAT: 11/04/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 15/04/13: Se agregarón las firmas en el pié de página.
'-------------------------------------------------------------------------------------------'
