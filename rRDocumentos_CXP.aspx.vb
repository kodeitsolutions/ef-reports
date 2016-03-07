'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDocumentos_CXP"
'-------------------------------------------------------------------------------------------'
Partial Class rRDocumentos_CXP
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'N/CR' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Nota_Credito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'N/DB' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Nota_Debito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'CHEQ' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Cheques_Devueltos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ISLR' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Retenciones_ISRL,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ADEL' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Adelentos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'GIRO' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Giros,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'CFXG' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Cambios_Facturas_Giros,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ATD' OR Cod_Tip = 'AJPA' OR Cod_Tip = 'AJPM' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Ajustes_positivos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ATC' OR Cod_Tip = 'AJNA' OR Cod_Tip = 'AJNM' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Ajustes_Negativos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'RETIVA' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Retencion_Impuesto,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'RETPAT' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Retencion_Patente,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'FACT' THEN Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Facturas,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'N/CR' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Nota_Credito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'N/DB' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Nota_Debito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'CHEQ' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Cheques_Devueltos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ISLR' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Retenciones_ISRL,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ADEL' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Adelentos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'GIRO' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Giros,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'CFXG' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Cambios_Facturas_Giros,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ATD' OR Cod_Tip = 'AJPA' OR Cod_Tip = 'AJPM' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Ajustes_positivos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'ATC' OR Cod_Tip = 'AJNA' OR Cod_Tip = 'AJNM' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Ajustes_Negativos,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'RETIVA' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Retencion_Impuesto,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'RETPAT' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Retencion_Patente,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cod_Tip = 'FACT' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Facturas")
            loComandoSeleccionar.AppendLine("INTO	#Temp")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("WHERE	Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)


            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine(" 		SUM(Nota_Credito) AS Nota_Credito,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Nota_Debito) AS Nota_Debito,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Cheques_Devueltos) AS Cheques_Devueltos,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Retenciones_ISRL) AS Retenciones_ISRL,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Adelentos) AS Adelentos,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Giros) AS Giros,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Cambios_Facturas_Giros) AS Cambios_Facturas_Giros, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Ajustes_positivos) AS Ajustes_positivos,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Ajustes_Negativos) AS Ajustes_Negativos,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Retencion_Impuesto) AS Retencion_Impuesto,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Retencion_Patente) AS Retencion_Patente, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Facturas) AS Facturas,")
            loComandoSeleccionar.AppendLine(" 		(SUM(Nota_Credito) + SUM(Nota_Debito) + SUM(Cheques_Devueltos) + SUM(Retenciones_ISRL) + SUM(Adelentos) +")
            loComandoSeleccionar.AppendLine(" 			SUM(Giros) + SUM(Cambios_Facturas_Giros) + SUM(Ajustes_positivos) + SUM(Ajustes_Negativos) + SUM(Retencion_Impuesto) +")
            loComandoSeleccionar.AppendLine(" 			SUM(Retencion_Patente)) AS Total_Mon, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Nota_Credito) AS Doc_Nota_Credito,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Nota_Debito) AS Doc_Nota_Debito,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Cheques_Devueltos) AS Doc_Cheques_Devueltos,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Retenciones_ISRL) AS Doc_Retenciones_ISRL,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Adelentos) AS Doc_Adelentos,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Giros) AS Doc_Giros,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Cambios_Facturas_Giros) AS Doc_Cambios_Facturas_Giros,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Ajustes_positivos) AS Doc_Ajustes_positivos,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Ajustes_Negativos) AS Doc_Ajustes_Negativos,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Retencion_Impuesto) AS Doc_Retencion_Impuesto,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Retencion_Patente) AS Doc_Retencion_Patente,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Facturas) AS Doc_Facturas,")
            loComandoSeleccionar.AppendLine(" 		(SUM(Doc_Nota_Credito) + SUM(Doc_Nota_Debito) + SUM(Doc_Cheques_Devueltos) + SUM(Doc_Retenciones_ISRL) +")
            loComandoSeleccionar.AppendLine(" 			SUM(Doc_Adelentos) + SUM(Doc_Giros) + SUM(Doc_Cambios_Facturas_Giros) + SUM(Doc_Ajustes_positivos) +")
            loComandoSeleccionar.AppendLine(" 			SUM(Doc_Ajustes_Negativos) + SUM(Doc_Retencion_Impuesto) + SUM(Doc_Retencion_Patente)) AS Total_Doc, ")
            loComandoSeleccionar.AppendLine(" 		1 AS Tabla")
			loComandoSeleccionar.AppendLine("INTO	#temp3")
			loComandoSeleccionar.AppendLine("FROM	#temp")

            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Cheque' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Cheque,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Deposito' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Deposito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Efectivo' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Efectivo,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Tarjeta' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Tarjeta,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Ticket' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Ticket,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Transferencia' THEN Detalles_Pagos.Mon_Net")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Transferencia,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Cheque' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Cheque,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Deposito' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Deposito,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Efectivo' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Efectivo,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Tarjeta' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Tarjeta,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Ticket' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Ticket,")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Tip_Ope = 'Transferencia' THEN 1")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END Doc_Transferencia")
            loComandoSeleccionar.AppendLine("INTO	#Temp2")
            loComandoSeleccionar.AppendLine("FROM	Detalles_Pagos")
            loComandoSeleccionar.AppendLine("	JOIN Pagos ON  Pagos.Documento = Detalles_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("WHERE	Pagos.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Pro between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Mon between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)

            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine(" 		SUM(Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Deposito) AS Deposito,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine(" 		SUM(Transferencia) AS Transferencia,  ")
            loComandoSeleccionar.AppendLine(" 		(SUM(Cheque) + SUM(Deposito) + SUM(Efectivo) + SUM(Tarjeta) + SUM(Ticket) + SUM(Transferencia)) AS Total_Mon_Cob,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Cheque) AS Doc_Cheque,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Deposito) AS Doc_Deposito,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Efectivo) AS Doc_Efectivo,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Tarjeta) AS Doc_Tarjeta,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Ticket) AS Doc_Ticket,")
            loComandoSeleccionar.AppendLine(" 		SUM(Doc_Transferencia) AS Doc_Transferencia,")
            loComandoSeleccionar.AppendLine(" 		(SUM(Doc_Cheque) + SUM(Doc_Deposito) + SUM(Doc_Efectivo) + SUM(Doc_Tarjeta) +  SUM(Doc_Ticket) + SUM(Doc_Transferencia)) AS Total_Doc_Cob, ")
            loComandoSeleccionar.AppendLine(" 		1 AS Tabla")
            loComandoSeleccionar.AppendLine("INTO 	#temp4")
            loComandoSeleccionar.AppendLine("FROM 	#temp2")

            loComandoSeleccionar.AppendLine("SELECT * ")
            loComandoSeleccionar.AppendLine("FROM #temp3, #temp4 WHERE #temp3.tabla = #temp4.tabla ")

            'loComandoSeleccionar.AppendLine(" SELECT * ")
            'loComandoSeleccionar.AppendLine(" FROM #temp10 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRDocumentos_CXP", laDatosReporte)

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
			
			
		'Formato de fecha y hora de impresión (manual)
			Dim loControl As CrystalDecisions.CrystalReports.Engine.TextObject  
			
			loControl = loObjetoReporte.ReportDefinition.ReportObjects("txtFechaImpresion")  
			loControl.Text = goServicios.mObtenerFormatoCadena(Date.Now(), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
			loControl = loObjetoReporte.ReportDefinition.ReportObjects("txtHoraImpresion")  
			loControl.Text = Strings.Format(Date.Now(), "hh:mm:ss tt")
			
			
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrRDocumentos_CXP.ReportSource = loObjetoReporte

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
' CMS: 24/09/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 18/04/12: Ajuste de formato de feha de impresión para que la tome del usuario.		'
'-------------------------------------------------------------------------------------------'
