Imports System.Data
Partial Class fOrdenes_Cheque_BOD_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Bancos.nom_Ban,'') AS nom_Ban, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Detalles_oPagos.num_doc,'') as cheque, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif,  ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) END) END) AS  Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE  ")
            loComandoSeleccionar.AppendLine("              (CASE WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Pagos.Telefonos END) END) AS  Telefonos,  ")
            loComandoSeleccionar.AppendLine("          Proveedores.Fax,  ")
            loComandoSeleccionar.AppendLine("		   CAST(DAY(Ordenes_Pagos.Fec_Ini) AS VARCHAR) +' de ' +   ")
            loComandoSeleccionar.AppendLine("		   (CASE MONTH(Ordenes_Pagos.Fec_Ini)   ")
            loComandoSeleccionar.AppendLine("				WHEN 1 THEN 'Enero'   ")
            loComandoSeleccionar.AppendLine("				WHEN 2 THEN 'Febrero'   ")
            loComandoSeleccionar.AppendLine("				WHEN 3 THEN 'Marzo'   ")
            loComandoSeleccionar.AppendLine("				WHEN 4 THEN 'Abril'   ")
            loComandoSeleccionar.AppendLine("				WHEN 5 THEN 'Mayo'   ")
            loComandoSeleccionar.AppendLine("				WHEN 6 THEN 'Junio'   ")
            loComandoSeleccionar.AppendLine("				WHEN 7 THEN 'Julio'   ")
            loComandoSeleccionar.AppendLine("				WHEN 8 THEN 'Agosto'   ")
            loComandoSeleccionar.AppendLine("				WHEN 9 THEN 'Septiembre'   ")
            loComandoSeleccionar.AppendLine("				WHEN 10 THEN 'Octubre'   ")
            loComandoSeleccionar.AppendLine("				WHEN 11 THEN 'Noviembre'   ")
            loComandoSeleccionar.AppendLine("				ELSE 'Diciembre'   ")
            loComandoSeleccionar.AppendLine("			END) as Mes_Dia,   ")
            loComandoSeleccionar.AppendLine("			CAST(YEAR(Ordenes_Pagos.Fec_Ini) AS VARCHAR) as año,   ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nom_Pro           As  Nombre_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Rif               As  Rif_Genenerico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nit               As  Nit_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Dir_Fis           As  Dir_Fis_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Telefonos         As  Telefonos_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento         As  Documento,  ")
            loComandoSeleccionar.AppendLine("          Ordenes_Pagos.Fec_Ini           As  Fec_ini,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Fin           As  Fec_Fin,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru           As  Mon_Bru_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp           As  Mon_Imp1_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net           As  Mon_Net_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret           As  Mon_Ret_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo            As  Motivo,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con        As  Cod_Con,  ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon        As  Renglon,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb        As  Mon_Deb,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab        As  Mon_Hab,  ")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net * -1  ")
            loComandoSeleccionar.AppendLine("            	ELSE  ")
            loComandoSeleccionar.AppendLine("                    Renglones_oPagos.Mon_Net ")
            loComandoSeleccionar.AppendLine("            END       As  Mon_Net_Ren,  ")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN  ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1  ")
            loComandoSeleccionar.AppendLine("            	ELSE  ")
            loComandoSeleccionar.AppendLine("                    Renglones_oPagos.Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("            END       As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Imp        As  Cod_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Comentario     As  Comentario_Ren,  ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1       As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))           As  Mon_Let  ")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN	Detalles_oPagos ON Detalles_opagos.documento = Ordenes_Pagos.documento ")
            loComandoSeleccionar.AppendLine("			AND Detalles_oPagos.renglon = '1'  ")
            loComandoSeleccionar.AppendLine("			AND Detalles_oPagos.tip_ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Bancarias ON Detalles_oPagos.cod_cue = Cuentas_Bancarias.cod_cue ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ON Bancos.cod_ban = Cuentas_Bancarias.cod_ban ")
            loComandoSeleccionar.AppendLine("		JOIN  Renglones_oPagos ON Ordenes_Pagos.Documento =   Renglones_oPagos.Documento ")
            loComandoSeleccionar.AppendLine("		JOIN  Proveedores	ON Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro  ")
            loComandoSeleccionar.AppendLine("        JOIN  Conceptos ON Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("            UNION ALL ")

            loComandoSeleccionar.AppendLine(" SELECT Ordenes_Pagos.Cod_Pro,  ")
            loComandoSeleccionar.AppendLine("			COALESCE(Bancos.nom_Ban,'') AS nom_Ban, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Detalles_oPagos.num_doc,'') as cheque, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif,  ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) END) END) AS  Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE  ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Pagos.Telefonos END) END) AS  Telefonos,  ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax,  ")
            loComandoSeleccionar.AppendLine("		   CAST(DAY(Ordenes_Pagos.Fec_Ini) AS VARCHAR) +' de ' +  ")
            loComandoSeleccionar.AppendLine("		   (CASE MONTH(Ordenes_Pagos.Fec_Ini)  ")
            loComandoSeleccionar.AppendLine("				WHEN 1 THEN 'Enero'  ")
            loComandoSeleccionar.AppendLine("				WHEN 2 THEN 'Febrero'  ")
            loComandoSeleccionar.AppendLine("				WHEN 3 THEN 'Marzo'  ")
            loComandoSeleccionar.AppendLine("				WHEN 4 THEN 'Abril'  ")
            loComandoSeleccionar.AppendLine("				WHEN 5 THEN 'Mayo'  ")
            loComandoSeleccionar.AppendLine("				WHEN 6 THEN 'Junio'  ")
            loComandoSeleccionar.AppendLine("				WHEN 7 THEN 'Julio'  ")
            loComandoSeleccionar.AppendLine("				WHEN 8 THEN 'Agosto'  ")
            loComandoSeleccionar.AppendLine("				WHEN 9 THEN 'Septiembre'  ")
            loComandoSeleccionar.AppendLine("				WHEN 10 THEN 'Octubre'  ")
            loComandoSeleccionar.AppendLine("				WHEN 11 THEN 'Noviembre'  ")
            loComandoSeleccionar.AppendLine("				ELSE 'Diciembre'  ")
            loComandoSeleccionar.AppendLine("			END) as Mes_Dia,  ")
            loComandoSeleccionar.AppendLine("			CAST(YEAR(Ordenes_Pagos.Fec_Ini) AS VARCHAR) as año,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nom_Pro           As  Nombre_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Rif               As  Rif_Genenerico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nit               As  Nit_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Dir_Fis           As  Dir_Fis_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Telefonos         As  Telefonos_Generico,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento         As  Documento,  ")
            loComandoSeleccionar.AppendLine("          Ordenes_Pagos.Fec_Ini           As  Fec_ini,  ")
            loComandoSeleccionar.AppendLine("         Ordenes_Pagos.Fec_Fin           As  Fec_Fin,  ")
            loComandoSeleccionar.AppendLine("          Ordenes_Pagos.Mon_Bru           As  Mon_Bru_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp           As  Mon_Imp1_Enc,  ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net           As  Mon_Net_Enc,  ")
            loComandoSeleccionar.AppendLine("          Ordenes_Pagos.Mon_Ret           As  Mon_Ret_Enc,  ")
            loComandoSeleccionar.AppendLine("          Ordenes_Pagos.Motivo            As  Motivo,  ")
            loComandoSeleccionar.AppendLine("                'ISLR'        As  Cod_Con,  ")
            loComandoSeleccionar.AppendLine("                'Retención de ISLR'    As  Nom_Con,  ")
            loComandoSeleccionar.AppendLine("          retenciones_documentos.Renglon        As  Renglon,  ")
            loComandoSeleccionar.AppendLine("           CAST(0 AS DECIMAL(28,10))      As  Mon_Deb,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Mon_RET        As  Mon_Hab,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Mon_RET*-1       As  Mon_Net_Ren,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Mon_Imp        As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Cod_Imp        As  Cod_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Comentario     As  Comentario_Ren,  ")
            loComandoSeleccionar.AppendLine("           retenciones_documentos.Mon_Imp       As  Mon_Imp_Ren,  ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))           As  Mon_Let ")
            loComandoSeleccionar.AppendLine(" FROM Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN	Detalles_oPagos ON Detalles_opagos.documento = Ordenes_Pagos.documento ")
            loComandoSeleccionar.AppendLine("			AND Detalles_oPagos.renglon = '1'  ")
            loComandoSeleccionar.AppendLine("			AND Detalles_oPagos.tip_ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Bancarias ON Detalles_oPagos.cod_cue = Cuentas_Bancarias.cod_cue ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ON Bancos.cod_ban = Cuentas_Bancarias.cod_ban ")
            loComandoSeleccionar.AppendLine(" 	JOIN  retenciones_documentos ON retenciones_documentos.Documento =   Ordenes_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("		AND retenciones_documentos.origen = 'Ordenes_Pagos'  ")
            loComandoSeleccionar.AppendLine("	JOIN  Proveedores	ON Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro  ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")



            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Cheque_BOD_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Cheque_BOD_GPV.ReportSource = loObjetoReporte

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
' EAG: 10/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
' EAG: 17/09/15: Se agregó a información de las retenciones
'-------------------------------------------------------------------------------------------'
