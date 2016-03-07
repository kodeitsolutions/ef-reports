Imports System.Data
Partial Class fPagos_Cheque02
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

           	loComandoSeleccionar.AppendLine("SELECT	    Pagos.Cod_Pro,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Rif,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Nit,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,") 
			loComandoSeleccionar.AppendLine("           Proveedores.Fax,") 
			loComandoSeleccionar.AppendLine("           Pagos.Documento,") 
			loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini,") 
			loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin,") 
			loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru			  As  Mon_Bru_Enc,") 
			loComandoSeleccionar.AppendLine("           (Pagos.Mon_Des * -1)			  As  Mon_Des,")
			loComandoSeleccionar.AppendLine("           Pagos.Mon_Net			  As  Mon_Net_Enc,") 
			loComandoSeleccionar.AppendLine("           (Pagos.Mon_Ret * -1)			  As  Mon_Ret_Enc,") 
			loComandoSeleccionar.AppendLine("           Pagos.Comentario			  As  Comentario,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Doc_Ori,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos.Renglon,") 
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Bru    As  Mon_Bru,")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Imp    As  Mon_Imp,") 
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo    As  Mon_Abo,") 
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net    As  Mon_Net_Ren,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Bru ELSE (Renglones_Pagos.Mon_Bru * -1) END)  AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Imp ELSE (Renglones_Pagos.Mon_Imp * -1) END)  AS  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Abo ELSE (Renglones_Pagos.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Net ELSE (Renglones_Pagos.Mon_Net * -1) END)  AS  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))       As  Mon_Let ")
			loComandoSeleccionar.AppendLine(" FROM      Pagos,") 
			loComandoSeleccionar.AppendLine("           Renglones_Pagos,") 
			loComandoSeleccionar.AppendLine("           Proveedores ")
			loComandoSeleccionar.AppendLine(" WHERE     Pagos.Documento =   Renglones_Pagos.Documento AND") 
			loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro   =   Proveedores.Cod_Pro ") 
			loComandoSeleccionar.AppendLine("		    AND" & cusAplicacion.goFormatos.pcCondicionPrincipal )         
		  
            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


			loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Cheque02", laDatosReporte)
			'loObjetoReporte.PrintOptions.PaperOrientation=CrystalDecisions.Shared.PaperOrientation.Landscape 
			'Dim i As New CrystalDecisions.Shared.PaperSize()
			
						
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPagos_Cheque02.ReportSource = loObjetoReporte

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
' GCR: 31/03/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 31/08/09: se multiplico -1 el monto bruto. impuesto, neto y el abonado segun su 
'                naturaleza
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Asignacion del signo negativo de la Retencion
'-------------------------------------------------------------------------------------------'
