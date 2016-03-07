Imports System.Data
Partial Class fOrdenes_Pagos_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Pagos.Telefonos END) END) AS  Telefonos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nom_Pro           As  Nombre_Generico, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Rif               As  Rif_Genenerico, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nit               As  Nit_Generico, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Dir_Fis           As  Dir_Fis_Generico, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Telefonos         As  Telefonos_Generico, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento         As  Documento, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini           As  Fec_ini, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Fin           As  Fec_Fin, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru           As  Mon_Bru_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp           As  Mon_Imp1_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net           As  Mon_Net_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret           As  Mon_Ret_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo            As  Motivo, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con        As  Cod_Con, ")
            'loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon        As  Renglon, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb        As  Mon_Deb, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab        As  Mon_Hab, ")
            ''loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Net        As  Mon_Net_Ren, ")
            'loComandoSeleccionar.AppendLine("            CASE ")
            'loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            'loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net * -1 ")
            'loComandoSeleccionar.AppendLine("            	ELSE ")
            'loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net ")
            'loComandoSeleccionar.AppendLine("            END       As  Mon_Net_Ren, ")
            ''loComandoSeleccionar.AppendLine("           Renglones_oPagos.Por_Imp1       As  Por_Imp_Ren, ")
            'loComandoSeleccionar.AppendLine("            CASE ")
            'loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            'loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1 ")
            'loComandoSeleccionar.AppendLine("            	ELSE ")
            'loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 ")
            'loComandoSeleccionar.AppendLine("            END       As  Mon_Imp_Ren, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Imp        As  Cod_Imp_Ren, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Comentario     As  Comentario_Ren, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1       As  Mon_Imp_Ren, ")
            'loComandoSeleccionar.AppendLine("           CAST('' as char(400))           As  Mon_Let ")
            'loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, ")
            'loComandoSeleccionar.AppendLine("           Renglones_oPagos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores, ")
            'loComandoSeleccionar.AppendLine("           Conceptos ")
            'loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Documento =   Renglones_oPagos.Documento AND ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro AND ")
            'loComandoSeleccionar.AppendLine("           Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro,  ")
            loComandoSeleccionar.AppendLine("			'**'+RTRIM(Proveedores.Nom_Pro)+'**' AS Nom_Pro,  ")
            loComandoSeleccionar.AppendLine("			RTRIM(Proveedores.Nom_Pro)  AS f,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Fax,  ")
            loComandoSeleccionar.AppendLine("			(CASE  ")
            loComandoSeleccionar.AppendLine("				WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro  ")
            loComandoSeleccionar.AppendLine("				ELSE (CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro  ")
            loComandoSeleccionar.AppendLine("						ELSE Ordenes_Pagos.Nom_Pro  ")
            loComandoSeleccionar.AppendLine("					  END)  ")
            loComandoSeleccionar.AppendLine("			END)																							AS  Nom_Pro,  ")
            loComandoSeleccionar.AppendLine("			(CASE ")
            loComandoSeleccionar.AppendLine("				WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif  ")
            loComandoSeleccionar.AppendLine("				ELSE (CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif  ")
            loComandoSeleccionar.AppendLine("						ELSE Ordenes_Pagos.Rif  ")
            loComandoSeleccionar.AppendLine("					 END)  ")
            loComandoSeleccionar.AppendLine("			END)																							AS  Rif,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit,  ")
            loComandoSeleccionar.AppendLine("			(CASE ")
            loComandoSeleccionar.AppendLine("				WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200)  ")
            loComandoSeleccionar.AppendLine("				ELSE (CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200)  ")
            loComandoSeleccionar.AppendLine("						ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200)  ")
            loComandoSeleccionar.AppendLine("					  END)  ")
            loComandoSeleccionar.AppendLine("			END)																							AS  Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("			(CASE ")
            loComandoSeleccionar.AppendLine("				WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos  ")
            loComandoSeleccionar.AppendLine("				ELSE (CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos  ")
            loComandoSeleccionar.AppendLine("						ELSE Ordenes_Pagos.Telefonos  ")
            loComandoSeleccionar.AppendLine("					  END)  ")
            loComandoSeleccionar.AppendLine("			END)																							AS  Telefonos,  ")
            loComandoSeleccionar.AppendLine("			Proveedores.Fax,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Nom_Pro																			As  Nombre_Generico,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Rif																				As  Rif_Genenerico,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Nit																				As  Nit_Generico,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Dir_Fis																			As  Dir_Fis_Generico,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Telefonos																			As  Telefonos_Generico,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Documento																			As  Documento,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini																			As  Fec_ini,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Fin																			As  Fec_Fin,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Bru																			As  Mon_Bru_Enc,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Imp																			As  Mon_Imp1_Enc,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net																			As  Mon_Net_Enc,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Ret																			As  Mon_Ret_Enc,  ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Motivo																			As  Motivo, ")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Renglon																			As  Renglon_Cheque, --Nuevo hasta ")
            loComandoSeleccionar.AppendLine("			COALESCE(Detalles_oPagos.Num_Doc,'[SIN CHEQUE]')																			As  Referencia,  ")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Fec_ini																			As  Fec_Ini_Cheque,  ")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Mon_Net																			As  Mon_Net_Cheque,  ")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Cod_Cue																			As  Cod_Cue_Che,  ")
            loComandoSeleccionar.AppendLine("			Cuentas_Bancarias.Nom_Cue																		As  Nom_Cue_Cheque,--Aqui  ")
            loComandoSeleccionar.AppendLine("			bancos.Nom_Ban																					As  Nom_Ban_Che,    ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Cod_Con																		As  Cod_Con,  ")
            loComandoSeleccionar.AppendLine("			Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)								As  Nom_Con,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Renglon																		As  Renglon,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Mon_Deb																		As  Mon_Deb,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Mon_Hab																		As  Mon_Hab,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Mon_Net																		As  Mon_Net_Ren,  ")
            loComandoSeleccionar.AppendLine("			(CASE  ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net * -1  ")
            loComandoSeleccionar.AppendLine("            	ELSE  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net  ")
            loComandoSeleccionar.AppendLine("            END)																							As  Mon_Net_Ren,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Por_Imp1       As  Por_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("			(CASE  ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1  ")
            loComandoSeleccionar.AppendLine("            	ELSE  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1  ")
            loComandoSeleccionar.AppendLine("            END)																							As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Cod_Imp																		As  Cod_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Comentario																		As  Comentario_Ren,  ")
            loComandoSeleccionar.AppendLine("			Renglones_oPagos.Mon_Imp1																		As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("			CAST('' as char(400))																			As  Mon_Let  ")
            loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_oPagos ON  Renglones_oPagos.Documento = Ordenes_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("   JOIN    Proveedores ON  Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("   JOIN	Conceptos ON Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Detalles_oPagos ON Detalles_oPagos.Documento = Ordenes_Pagos.Documento		--Nuevo hasta ")
            loComandoSeleccionar.AppendLine("		AND	Detalles_oPagos.tip_ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("		AND Detalles_oPagos.Renglon = '1'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue   =   Detalles_oPagos.Cod_Cue ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Bancos ON Bancos.Cod_ban   =   Cuentas_Bancarias.Cod_Ban					--aqui ")
            loComandoSeleccionar.AppendLine("WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")



            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = "**" & goServicios.mConvertirMontoLetras(lnMontoNumero) & "**"

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Pagos_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Pagos_GPV.ReportSource = loObjetoReporte

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
' EAG: 17/08/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
