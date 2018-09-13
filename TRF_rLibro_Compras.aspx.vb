'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rLibro_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rLibro_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @FecIni AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @FecFin	AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero As Decimal(28, 10) = CAST(0 As Decimal(28, 10));")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio As NVARCHAR(30) = N''; ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() OVER (ORDER BY Registros.Fecha_Doc) AS Num,*")
            loComandoSeleccionar.AppendLine("FROM(")
            loComandoSeleccionar.AppendLine("	/*Obtener facturas*/")
            loComandoSeleccionar.AppendLine("   SELECT  Cuentas_Pagar.Cod_Tip					AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento 				AS Documento,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori					AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini			        AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif					        AS RIF,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro				        AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura			        AS Num_Factura,")
            loComandoSeleccionar.AppendLine("			@lcVacio						        AS Num_Comprobante,")
            loComandoSeleccionar.AppendLine("			@lcVacio						        AS Factura_Afect,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Net			        AS Total_Doc,")
            loComandoSeleccionar.AppendLine("           CASE WHEN (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731') OR (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Reg >= '20180801')")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Net/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("           END                                     AS Total_Doc,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Exe			        AS Monto_Exento,")
            loComandoSeleccionar.AppendLine("           CASE WHEN (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731') OR (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Reg >= '20180801')")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Exe/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Exe")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Exento,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Bas1                  AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("           CASE WHEN (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731') OR (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Reg >= '20180801')")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Bas1/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Bas1")
            loComandoSeleccionar.AppendLine("           END                                     AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Imp1			        AS Impuesto,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1                  AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("           CASE WHEN (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731') OR (Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Reg >= '20180801')")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Imp1/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Imp1")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Monto_Ret,")
            loComandoSeleccionar.AppendLine("			@FecIni									AS Fecha_De,")
            loComandoSeleccionar.AppendLine("           @FecFin									AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("   FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("   	JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   WHERE Cuentas_Pagar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("   	AND Cuentas_Pagar.Fec_Reg BETWEEN @FecIni AND @FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("	UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   /*Obtener notas de crédito*/")
            loComandoSeleccionar.AppendLine("   SELECT  Cuentas_Pagar.Cod_Tip					AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Documento 				AS Documento,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Doc_Ori					AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini			        AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Rif					        AS RIF,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Nom_Pro				        AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Factura			        AS Num_Factura,")
            loComandoSeleccionar.AppendLine("   		@lcVacio						        AS Num_Comprobante,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Referencia		        AS Factura_Afect,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Net*(-1)			    AS Total_Doc,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Net/100000)*(-1)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Net*(-1)")
            loComandoSeleccionar.AppendLine("           END                                     AS Total_Doc,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Exe*(-1)			    AS Monto_Exento,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Exe/100000)*(-1)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Exe*(-1)")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Exento,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Bas1*(-1)             AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Bas1/100000)*(-1)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Bas1*(-1)")
            loComandoSeleccionar.AppendLine("           END                                     AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Imp1				    AS Impuesto,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1*(-1)             AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Imp1/100000)*(-1)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Imp1*(-1)")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("   		@lnCero									AS Monto_Ret,")
            loComandoSeleccionar.AppendLine("   		@FecIni									AS Fecha_De,")
            loComandoSeleccionar.AppendLine("           @FecFin									AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("   FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("   	JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   WHERE Cuentas_Pagar.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("   	AND Cuentas_Pagar.Fec_Reg BETWEEN @FecIni AND @FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   /*Obtener notas de débito*/")
            loComandoSeleccionar.AppendLine("   SELECT  Cuentas_Pagar.Cod_Tip					AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Documento 				AS Documento,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Doc_Ori					AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini			        AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Rif					        AS RIF,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Nom_Pro				        AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Factura			        AS Num_Factura,")
            loComandoSeleccionar.AppendLine("   		@lcVacio						        AS Num_Comprobante,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Referencia		        AS Factura_Afect,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Net			        AS Total_Doc,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Net/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("           END                                     AS Total_Doc,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Exe			        AS Monto_Exento,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Exe/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Exe")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Exento,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Bas1                  AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Bas1/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Bas1")
            loComandoSeleccionar.AppendLine("           END                                     AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Imp1			        AS Impuesto,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1                  AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Cuentas_Pagar.Mon_Imp1/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Cuentas_Pagar.Mon_Imp1")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Imp,")
            loComandoSeleccionar.AppendLine("   		@lnCero									AS Monto_Ret,")
            loComandoSeleccionar.AppendLine("   		@FecIni									AS Fecha_De,")
            loComandoSeleccionar.AppendLine("           @FecFin									AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("   FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("   	JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   WHERE Cuentas_Pagar.Cod_Tip = 'N/DB'")
            loComandoSeleccionar.AppendLine("   	AND Cuentas_Pagar.Fec_Reg BETWEEN @FecIni AND @FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("	UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   SELECT	Cuentas_Pagar.Cod_Tip					AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Documento 				AS Documento,")
            loComandoSeleccionar.AppendLine("   		Cuentas_Pagar.Doc_Ori					AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini			        AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Rif						    AS RIF,")
            loComandoSeleccionar.AppendLine("   		Proveedores.Nom_Pro					    AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("   		@lcVacio							    AS Num_Factura,")
            loComandoSeleccionar.AppendLine("           CASE WHEN MONTH(Cuentas_Pagar.Fec_Ini) < 10")
            loComandoSeleccionar.AppendLine("               THEN CONCAT(YEAR(Cuentas_Pagar.Fec_Ini),'0', MONTH(Cuentas_Pagar.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("               ELSE CONCAT(YEAR(Cuentas_Pagar.Fec_Ini), MONTH(Cuentas_Pagar.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("           END										AS Num_Comprobante,")
            loComandoSeleccionar.AppendLine("   		Documentos.Factura					    AS Factura_Afect,")
            'loComandoSeleccionar.AppendLine("			Documentos.Mon_Net					    AS Total_Doc,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Documentos.Mon_Net/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Documentos.Mon_Net")
            loComandoSeleccionar.AppendLine("           END                                     AS Total_Doc,")
            'loComandoSeleccionar.AppendLine("			Documentos.Mon_Exe					    AS Monto_Exento,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                THEN (Documentos.Mon_Exe/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Documentos.Mon_Exe")
            loComandoSeleccionar.AppendLine("           END                                     AS Monto_Exento,")
            'loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Bas		    AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("           CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820'")
            loComandoSeleccionar.AppendLine("                THEN (Retenciones_Documentos.Mon_Bas/100000)")
            loComandoSeleccionar.AppendLine("                ELSE Retenciones_Documentos.Mon_Bas")
            loComandoSeleccionar.AppendLine("           END                                     AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Por_Ret		    AS Impuesto, ")
            loComandoSeleccionar.AppendLine("			@lnCero                     		    AS Monto_Imp,")
            'loComandoSeleccionar.AppendLine("			CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            'loComandoSeleccionar.AppendLine("				 THEN Retenciones_Documentos.Mon_Ret * (-1)")
            'loComandoSeleccionar.AppendLine("				 ELSE Retenciones_Documentos.Mon_Ret")
            'loComandoSeleccionar.AppendLine("			END 									AS Monto_Ret,")
            loComandoSeleccionar.AppendLine("			CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("				 THEN (CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                           THEN (Retenciones_Documentos.Mon_Ret/100000) * (-1)")
            loComandoSeleccionar.AppendLine("                           ELSE Retenciones_Documentos.Mon_Ret * (-1) END)")
            loComandoSeleccionar.AppendLine("				 ELSE (CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini > '20180731'")
            loComandoSeleccionar.AppendLine("                           THEN (Retenciones_Documentos.Mon_Ret/100000)")
            loComandoSeleccionar.AppendLine("                           ELSE Retenciones_Documentos.Mon_Ret END)")
            loComandoSeleccionar.AppendLine("			END 									AS Monto_Ret,")
            loComandoSeleccionar.AppendLine("           @FecIni                                 AS Fecha_De,")
            loComandoSeleccionar.AppendLine("           @FecFin                                 AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("   FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("       JOIN Cuentas_Pagar AS Documentos ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("   	    AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("   	JOIN	Retenciones_Documentos ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("   		AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("   	JOIN	Proveedores ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("   	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("   WHERE Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("       AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("   	AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("   	AND Cuentas_Pagar.Fec_Ini BETWEEN @FecIni AND @FecFin) Registros ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rLibro_Compras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTRF_rLibro_Compras.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 06/10/14: Codigo inicial.					                                        '
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/15: Se eliminó el campo "Fecha de contabilización" y se agregó el número de    '
'                factura del proveedor.					                                    '
'-------------------------------------------------------------------------------------------'
' RJG: 29/04/15: Se ampliaron los campos de Factura y Control. Se ajustó el ordenamiento del'
'                grupo de ajustes. Se agregó el "impuesto excluido" de la GACETA 6152.		'
'-------------------------------------------------------------------------------------------'
