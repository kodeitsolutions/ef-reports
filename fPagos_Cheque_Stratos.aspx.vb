'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPagos_Cheque_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class fPagos_Cheque_Stratos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	Pagos.Cod_Pro,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Rif,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Nit,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Fax,") 
			loComandoSeleccionar.AppendLine("           Pagos.Documento,") 
            loComandoSeleccionar.AppendLine("           Pagos.Status,")
            loComandoSeleccionar.AppendLine(" 	CASE  ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 1 THEN 'Jan '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar)")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 2 THEN 'Feb '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 3 THEN 'Mar '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 4 THEN 'Apr '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 5 THEN 'May '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 6 THEN 'Jun '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 7 THEN 'Jul '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 8 THEN 'Aug '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 9 THEN 'Sep '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', '	+ Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 10 THEN 'Oct '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 11 THEN 'Nov '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pagos.Fec_Ini) = 12 THEN 'Dec '+ Cast(DATEPART(dd, Pagos.Fec_Ini) AS VArchar)+ ', ' + Cast(DATEPART(YYYY, Pagos.Fec_Ini) AS VArchar) ")
            loComandoSeleccionar.AppendLine(" 	END AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin,") 
			loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru			    As  Mon_Bru_Enc,") 
			loComandoSeleccionar.AppendLine("           (Pagos.Mon_Des * -1)        As  Mon_Des,") 
			loComandoSeleccionar.AppendLine("           Pagos.Mon_Net			    As  Mon_Net_Enc,") 
			loComandoSeleccionar.AppendLine("           (Pagos.Mon_Ret * -1)	    As  Mon_Ret_Enc,") 
			loComandoSeleccionar.AppendLine("           Pagos.Comentario		    As  Comentario,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip     As  Cod_Tip,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Doc_Ori     As  Doc_Ori,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Bru     As  Mon_Bru,")
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Imp     As  Mon_Imp,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo     As  Mon_Abo,") 
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net     As  Mon_Net_Ren,")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))       As  Mon_Let ,")
			loComandoSeleccionar.AppendLine("           Detalles_Pagos.Num_Doc,")
			loComandoSeleccionar.AppendLine("           ISNULL(Detalles_Pagos.Mon_Net,0) AS Mon_Net_Det")
			loComandoSeleccionar.AppendLine(" FROM      Pagos,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos,")
			loComandoSeleccionar.AppendLine("           Detalles_Pagos,")
			loComandoSeleccionar.AppendLine("           Proveedores ")
			loComandoSeleccionar.AppendLine(" WHERE     Pagos.Documento =   Renglones_Pagos.Documento AND")
			loComandoSeleccionar.AppendLine("           Pagos.Documento =   Detalles_Pagos.Documento AND")
			loComandoSeleccionar.AppendLine("           Detalles_Pagos.Tip_Ope NOT IN ('Efectivo') AND ")
			loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro   =   Proveedores.Cod_Pro ") 
			loComandoSeleccionar.AppendLine("		    AND" & cusAplicacion.goFormatos.pcCondicionPrincipal )         
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Det"))
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



            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Cheque_Stratos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPagos_Cheque_Stratos.ReportSource = loObjetoReporte

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
' JFP: 08/10/12: Adicion del status para la etiqueta
'-------------------------------------------------------------------------------------------'