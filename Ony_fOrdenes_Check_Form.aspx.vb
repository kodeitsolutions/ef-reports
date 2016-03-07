'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "Ony_fOrdenes_Check_Form"
'-------------------------------------------------------------------------------------------'
Partial Class Ony_fOrdenes_Check_Form
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine(" SELECT	Pagos.Cod_Pro,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Rif,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nit,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Fax,")
            'loComandoSeleccionar.AppendLine("           Pagos.Documento,")
            'loComandoSeleccionar.AppendLine(" 	CASE  ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 1 THEN 'Jan '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar)")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 2 THEN 'Feb '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 3 THEN 'Mar '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 4 THEN 'Apr '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 5 THEN 'May '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 6 THEN 'Jun '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 7 THEN 'Jul '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 8 THEN 'Aug '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 9 THEN 'Sep '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 10 THEN 'Oct '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 11 THEN 'Nov '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 12 THEN 'Dec '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            'loComandoSeleccionar.AppendLine(" 	END AS Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin,")
            'loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru			    As  Mon_Bru_Enc,")
            'loComandoSeleccionar.AppendLine("           (Pagos.Mon_Des * -1)        As  Mon_Des,")
            'loComandoSeleccionar.AppendLine("           Pagos.Mon_Net			    As  Mon_Net_Enc,")
            'loComandoSeleccionar.AppendLine("           (Pagos.Mon_Ret * -1)	    As  Mon_Ret_Enc,")
            'loComandoSeleccionar.AppendLine("           Pagos.Comentario		    As  Comentario,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip     As  Cod_Tip,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Doc_Ori     As  Doc_Ori,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Bru     As  Mon_Bru,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Imp     As  Mon_Imp,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo     As  Mon_Abo,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net     As  Mon_Net_Ren,")
            'loComandoSeleccionar.AppendLine("           CAST('' as char(400))       As  Mon_Let ,")
            'loComandoSeleccionar.AppendLine("           Detalles_Pagos.Num_Doc,")
            'loComandoSeleccionar.AppendLine("           ISNULL(Detalles_Pagos.Mon_Net,0) AS Mon_Net_Det")
            'loComandoSeleccionar.AppendLine(" FROM      Pagos,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos,")
            'loComandoSeleccionar.AppendLine("           Detalles_Pagos,")
            'loComandoSeleccionar.AppendLine("           Proveedores ")
            'loComandoSeleccionar.AppendLine(" WHERE     Pagos.Documento =   Renglones_Pagos.Documento AND")
            'loComandoSeleccionar.AppendLine("           Pagos.Documento =   Detalles_Pagos.Documento AND")
            'loComandoSeleccionar.AppendLine("           Detalles_Pagos.Tip_Ope NOT IN ('Efectivo') AND ")
            'loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro   =   Proveedores.Cod_Pro ")
            'loComandoSeleccionar.AppendLine("		    AND" & cusAplicacion.goFormatos.pcCondicionPrincipal)



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
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Comentario        As  Comentario, ")
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




            'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal

            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Cheque"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetrasEN(lnMontoNumero) & " Dollars"

            Next loFilas


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



            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("Ony_fOrdenes_Check_Form", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvOny_fOrdenes_Check_Form.ReportSource = loObjetoReporte

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
' CMS: 07/04/10: Codigo inicial
'-------------------------------------------------------------------------------------------'