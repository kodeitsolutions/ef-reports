Imports System.Data
Partial Class fPagos_Cheque_BOD_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Bancos.nom_Ban,'') AS Cod_Ban, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Detalles_Pagos.num_doc,'') as cheque, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("			Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("			CAST(DAY(Pagos.Fec_Ini) AS VARCHAR) +' de ' +  ")
            loComandoSeleccionar.AppendLine("			(CASE MONTH(Pagos.Fec_Ini)  ")
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
            loComandoSeleccionar.AppendLine("			CAST(YEAR(Pagos.Fec_Ini) AS VARCHAR) as año,  ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Bru			    As  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("			(Pagos.Mon_Des * -1)        As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Net			    As  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("			(Pagos.Mon_Ret * -1)	    As  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("			Pagos.Comentario		    As  Comentario, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Cod_Tip     As  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("                Case Renglones_Pagos.Factura ")
            loComandoSeleccionar.AppendLine("				WHEN '' THEN Renglones_Pagos.doc_ori ")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Pagos.Factura ")
            loComandoSeleccionar.AppendLine("			END     As  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("                    ''						   AS  Documento_Afectado, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Renglon     As  Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Bru     As  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Imp     As  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("			CASE Renglones_Pagos.tip_doc ")
            loComandoSeleccionar.AppendLine("					WHEN 'Debito' THEN Renglones_Pagos.Mon_Abo ")
            loComandoSeleccionar.AppendLine("					ELSE Renglones_Pagos.Mon_Abo* -1 ")
            loComandoSeleccionar.AppendLine("			END     As  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Net     As  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("			CAST('' as char(400))       As  Mon_Let  ")
            loComandoSeleccionar.AppendLine("FROM Pagos ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Detalles_Pagos ON Detalles_pagos.documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Detalles_Pagos.renglon = '1' ")
            loComandoSeleccionar.AppendLine("			AND Detalles_Pagos.tip_ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ON Bancos.cod_ban = Detalles_Pagos.cod_ban  ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Pagos ON Pagos.Documento =   Renglones_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("    JOIN   Proveedores ON Pagos.Cod_Pro   =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Bancos.nom_Ban,'') AS Cod_Ban, ")
            loComandoSeleccionar.AppendLine("			COALESCE(Detalles_Pagos.num_doc,'') as cheque, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("			Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("			CAST(DAY(Pagos.Fec_Ini) AS VARCHAR) +' de ' +  ")
            loComandoSeleccionar.AppendLine("			(CASE MONTH(Pagos.Fec_Ini)  ")
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
            loComandoSeleccionar.AppendLine("			CAST(YEAR(Pagos.Fec_Ini) AS VARCHAR) as año,  ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Bru					As  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("			(Pagos.Mon_Des * -1)				As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Net					As  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("			(Pagos.Mon_Ret * -1)				As  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("			Pagos.Comentario					As  Comentario, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.cla_des   As  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Doc_des   As  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			RTRIM(CASE Cuentas_Pagar.factura ")
            loComandoSeleccionar.AppendLine("				WHEN '' THEN Cuentas_Pagar.documento ")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.factura ")
            loComandoSeleccionar.AppendLine("			END)	+'/'+ RTRIM(Cuentas_Pagar.cod_tip)					AS  Documento_Afectado, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Renglon     As  Renglon, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Mon_Bru     As  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Mon_Imp     As  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Mon_RET*-1     As  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			retenciones_documentos.Mon_Net     As  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("			CAST('' as char(400))       As  Mon_Let  ")
            loComandoSeleccionar.AppendLine("FROM pagos  ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Detalles_Pagos ON Detalles_pagos.documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Detalles_Pagos.renglon = '1' ")
            loComandoSeleccionar.AppendLine("			AND Detalles_Pagos.tip_ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ON Bancos.cod_ban = Detalles_Pagos.cod_ban  ")
            loComandoSeleccionar.AppendLine("	JOIN retenciones_documentos ON retenciones_documentos.documento = Pagos.Documento ")
            loComandoSeleccionar.AppendLine("		AND retenciones_documentos.origen = 'pagos' ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.documento = retenciones_documentos.doc_ori ")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.cod_tip = retenciones_documentos.cla_ori ")
            loComandoSeleccionar.AppendLine("	JOIN   Proveedores ON Pagos.Cod_Pro   =   Proveedores.Cod_Pro  ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Cheque_BOD_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPagos_Cheque_BOD_GPV.ReportSource = loObjetoReporte

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
' EAG: 16/09/15: Agregar las retenciones al select a través de un union all
'-------------------------------------------------------------------------------------------'
