'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System
Imports System.Data
Imports System.Collections.Specialized
Imports System.Net
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_USA"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_USA
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			loComandoSeleccionar.AppendLine("SET @lnCero = 0")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		'Cobros'											AS Tabla, ")
			loComandoSeleccionar.AppendLine("			CAST(0 As INT)										As Contador, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                              AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Cobros.Fec_Ini                                      AS FechaCobro, ")
			loComandoSeleccionar.AppendLine("			Clientes.Rif										AS Rif, ")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli									AS Nom_Cli, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento							AS Documento, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control								AS Control,")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Doc_Ori								AS Doc_Ori,")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net                            	AS Mon_Net, ")
			loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) 	AS Mon_Bru, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Exe                            	AS Mon_Exe, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1                           	AS Mon_Imp1, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp		                        AS Cod_Imp, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1								AS Por_Imp1, ")
			loComandoSeleccionar.AppendLine("			'01-Reg'											AS Status_Documento, ")
			loComandoSeleccionar.AppendLine("			'01'												AS Tipo_Documento, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Status								AS Status, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal1                              AS Fiscal1, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal2                              AS Fiscal2, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal3                              AS Fiscal3, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal4                              AS Fiscal4, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal5                              AS Fiscal5, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des                              AS Mon_Des, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec                              AS Mon_Rec, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Des                              AS Por_Des, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Rec                              AS Por_Rec, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1                             AS Mon_Otr1, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2                             AS Mon_Otr2, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3                             AS Mon_Otr3, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Referencia                           AS Referencia, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Dis_Imp								AS Dis_Imp ")
			loComandoSeleccionar.AppendLine("INTO		#tmpLibroVentas ")
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
			loComandoSeleccionar.AppendLine("	JOIN 	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
			loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Cobros ON Cuentas_Cobrar.Documento	=   Renglones_Cobros.Doc_Ori")
			loComandoSeleccionar.AppendLine("							AND Renglones_Cobros.Cod_Tip	=   'RETIVA' ")
			loComandoSeleccionar.AppendLine("							AND Cuentas_Cobrar.Cod_Tip		=   Renglones_Cobros.Cod_Tip ")
			loComandoSeleccionar.AppendLine("	JOIN Cobros ON Cobros.Documento	=   Renglones_Cobros.Documento")
			loComandoSeleccionar.AppendLine("			AND Cobros.Fec_Ini		BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini      < " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status      <> 'Anulado' ")
			loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
			If lcParametro5Desde = "Igual" Then
			    loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
			Else
			    loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
			End If
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine(" UNION ALL ")


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT    'Cobros'												AS Tabla, ")
            loComandoSeleccionar.AppendLine("			CAST(0 As INT)										As Contador, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS FechaCobro, ")
            loComandoSeleccionar.AppendLine("			Clientes.Rif										As Rif,")
            loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli									AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento							AS Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control								AS Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Factura								AS Factura, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net								AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)	AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Exe								AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1								AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp			            		AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1								AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			'01-Reg'											AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			'02'												AS Tipo_Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Status								AS Status, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal1                              AS Fiscal1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal2                              AS Fiscal2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal3                              AS Fiscal3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal4                              AS Fiscal4, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal5                              AS Fiscal5, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des                              AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec                              AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Des                              AS Por_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Rec                              AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1                             AS Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2                             AS Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3                             AS Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Referencia                           AS Referencia, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Dis_Imp								AS Dis_Imp ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip     =   'FACT' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status <>   'Anulado' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine(" UNION ALL ")



            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		'Cobros'												AS Tabla, ")
            loComandoSeleccionar.AppendLine("			CAST(0 As INT)											AS Contador, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                	AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                	AS FechaCobro, ")
            loComandoSeleccionar.AppendLine("			Clientes.Rif											AS Rif, ")
            loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli										AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip									AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento								AS Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control									AS Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Factura									AS Factura, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)		AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Exe									AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1									AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp									AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1									AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			'01-Reg'												AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			'02'													AS Tipo_Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Status									AS Status, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal1        							AS Fiscal1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal2        							AS Fiscal2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal3        							AS Fiscal3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal4        							AS Fiscal4, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal5        							AS Fiscal5, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Des * 1) 							AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Rec	* 1) 							AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Des									AS Por_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Rec									AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr1 * 1)							AS Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr2 * 1)							AS Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr3 * 1)							AS Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Referencia     							AS Referencia, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Dis_Imp									AS Dis_Imp ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status <>   'Anulado' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine(" UNION ALL ")


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT    'Cobros' AS Tabla			, ")
            loComandoSeleccionar.AppendLine("			CAST(0 As INT)	As Contador	, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                      		AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                      		AS FechaCobro, ")
            loComandoSeleccionar.AppendLine("			Clientes.Rif														AS Rif, ")
            loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli													AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip												AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento											AS Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control												AS Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Factura												AS Factura, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net												AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)					AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Exe                        						AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1                       						AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp												AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1												AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			'01-Reg'															AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			'02'																AS Tipo_Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Status												AS Status, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal1                                              AS Fiscal1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal2                                              AS Fiscal2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal3                                              AS Fiscal3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal4                                              AS Fiscal4, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fiscal5                                              AS Fiscal5, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des                                              AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec                                              AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Des                                              AS Por_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Rec                                              AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1                                             AS Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2                                             AS Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3                                             AS Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Referencia                                           AS Referencia, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Dis_Imp												AS Dis_Imp ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Coloca los contadores de documentos antes de distribuir")
            loComandoSeleccionar.AppendLine("UPDATE		#tmpLibroVentas			")
            loComandoSeleccionar.AppendLine("SET		Contador = C.Contador")
            loComandoSeleccionar.AppendLine("FROM		(SELECT		ROW_NUMBER() OVER(	ORDER BY	Tabla ASC,")
            loComandoSeleccionar.AppendLine("									(CASE Cod_Tip")
            loComandoSeleccionar.AppendLine("										WHEN 'FACT'		THEN 1")
            loComandoSeleccionar.AppendLine("										WHEN 'RETIVA'	THEN 2")
            loComandoSeleccionar.AppendLine("										WHEN 'N/CR'		THEN 3")
            loComandoSeleccionar.AppendLine("										WHEN 'N/DB'		THEN 4")
            loComandoSeleccionar.AppendLine("										ELSE				 5")
            loComandoSeleccionar.AppendLine("									END) ASC, Fec_Ini ASC, ")
            loComandoSeleccionar.AppendLine("									Documento, Cod_Tip")
            loComandoSeleccionar.AppendLine("								) AS Contador,")
            loComandoSeleccionar.AppendLine("						Documento,")
            loComandoSeleccionar.AppendLine("						Cod_Tip")
            loComandoSeleccionar.AppendLine("				FROM	#tmpLibroVentas")
            loComandoSeleccionar.AppendLine("			) AS C")
            loComandoSeleccionar.AppendLine("WHERE	C.Documento = #tmpLibroVentas.Documento")
            loComandoSeleccionar.AppendLine("	AND	C.Cod_Tip = #tmpLibroVentas.Cod_Tip")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
           
            
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("-- AÑADE LOS IMPUESTOS DETALLADOS EXTRAIDOS DE LA DISTRIBUCIÓN ")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Cod_Tip, Documento, CAST(Dis_Imp AS XML) AS Impuestos")
            loComandoSeleccionar.AppendLine("INTO		#tmpImpuestos")
            loComandoSeleccionar.AppendLine("FROM		#tmpLibroVentas ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tmpImpuestos.Cod_Tip, #tmpImpuestos.Documento, ")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./codigo)[1]',		'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'')	AS Codigo,	")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./porcentaje)[1]',	'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'') 	AS porcentaje,")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./base)[1]',		'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'') 	AS base,	")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./exento)[1]',		'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'') 	AS exento,	")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./monto)[1]',		'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'') 	AS monto")
            loComandoSeleccionar.AppendLine("INTO	#tmpImpuestos2")
            loComandoSeleccionar.AppendLine("FROM	#tmpImpuestos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY #tmpImpuestos.Impuestos.nodes('/impuestos/impuesto') AS T(C)")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Tip, Documento, Porcentaje DESC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpImpuestos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Indexa los impuestos")
            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(PARTITION BY Cod_Tip, Documento ORDER BY Cod_Tip ASC, Documento ASC, Porcentaje DESC) AS Indice,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestos2.Cod_Tip, #tmpImpuestos2.Documento, #tmpImpuestos2.Codigo, ")
            loComandoSeleccionar.AppendLine("			CAST(#tmpImpuestos2.Porcentaje AS DECIMAL(25,10)) AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			CAST(#tmpImpuestos2.base	AS DECIMAL(25,10)) AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			CAST(#tmpImpuestos2.exento	AS DECIMAL(25,10)) AS Mon_Exe,	")
            loComandoSeleccionar.AppendLine("			CAST(#tmpImpuestos2.monto	AS DECIMAL(25,10)) AS Mon_Imp1	 ")
            loComandoSeleccionar.AppendLine("INTO		#tmpImpuestos3")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestos2")
            loComandoSeleccionar.AppendLine("ORDER BY	Cod_Tip ASC, Documento ASC, Porcentaje DESC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE	#tmpImpuestos2")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Organiza las columnas, separando los grabables de los exentos")
            loComandoSeleccionar.AppendLine("SELECT		Grabables.Indice						AS Indice,")
            loComandoSeleccionar.AppendLine("			Grabables.Cod_Tip						AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			Grabables.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("			Grabables.Codigo						AS Codigo,")
            loComandoSeleccionar.AppendLine("			Grabables.Porcentaje					AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			ISNULL(NoGrabables.Mon_Exe,	@lnCero)	AS Mon_Exe,")
            loComandoSeleccionar.AppendLine("			Grabables.Mon_Bru						AS Mon_Bru,")
            loComandoSeleccionar.AppendLine("			Grabables.Mon_Imp1						AS Mon_Imp1")
            loComandoSeleccionar.AppendLine("INTO		#tmpImpuestosFinal")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestos3 AS Grabables")
            loComandoSeleccionar.AppendLine("	LEFT JOIN		#tmpImpuestos3 AS NoGrabables")
            loComandoSeleccionar.AppendLine("		ON	NoGrabables.Documento	= Grabables.Documento")
            loComandoSeleccionar.AppendLine("		AND	NoGrabables.Cod_Tip		= Grabables.Cod_Tip  ")
            loComandoSeleccionar.AppendLine("		AND	NoGrabables.Porcentaje = 0")
            loComandoSeleccionar.AppendLine("		AND	Grabables.Porcentaje > 0")
            loComandoSeleccionar.AppendLine("WHERE		Grabables.Porcentaje > 0 OR NOGrabables.Indice = 1")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT		NoGrabables.Indice						AS Indice,")
            loComandoSeleccionar.AppendLine("			NoGrabables.Cod_Tip						AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			NoGrabables.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("			CAST('' AS NVARCHAR(10))				AS Codigo,")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			NoGrabables.Mon_Exe						AS Mon_Exe,")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Mon_Bru,")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Mon_Imp1")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestos3 AS NoGrabables")
            loComandoSeleccionar.AppendLine("WHERE		NoGrabables.Porcentaje = 0 ")
            loComandoSeleccionar.AppendLine("		AND	NOGrabables.Indice = 1")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpImpuestos3")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            
			loComandoSeleccionar.AppendLine("-- ********************************************************************")
			loComandoSeleccionar.AppendLine("-- EJECUTA EL SELECT FINAL ")
			loComandoSeleccionar.AppendLine("-- ********************************************************************")
			loComandoSeleccionar.AppendLine("--Organiza las columnas, separando los grabables de los exentos")
            loComandoSeleccionar.AppendLine("SELECT		(CASE WHEN #tmpImpuestosFinal.Indice = 1  ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Contador ELSE NULL END)			AS Contador	, ")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Indice									AS Indice	, ")
            loComandoSeleccionar.AppendLine("			#tmpLibroVentas.Tabla										AS Tabla	,")
            loComandoSeleccionar.AppendLine("			#tmpLibroVentas.Fec_Ini										AS Fec_Ini	,")
            loComandoSeleccionar.AppendLine("			#tmpLibroVentas.FechaCobro									AS FechaCobro,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Rif ELSE NULL END)					AS Rif		,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Nom_Cli ELSE NULL END)				AS Nom_Cli	,")
            loComandoSeleccionar.AppendLine("			#tmpLibroVentas.Cod_Tip										AS Cod_Tip	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Documento ELSE NULL END)			AS Documento,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Control ELSE NULL END) 			AS Control	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Doc_Ori ELSE NULL END) 			AS Doc_Ori	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Net ELSE NULL END) 			AS Mon_Net	,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Mon_Bru									AS Mon_Bru	,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Mon_Exe									AS Mon_Exe	,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Mon_Imp1									AS Mon_Imp1	,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Porcentaje								AS Por_Imp1	,")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Codigo									AS Cod_Imp	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Status_Documento ELSE NULL END)	AS Status_Documento	,")
            loComandoSeleccionar.AppendLine("			#tmpLibroVentas.Tipo_Documento								AS Tipo_Documento	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Status ELSE NULL END)				AS Status	,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Fiscal1 ELSE NULL END) 			AS Fiscal1	,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Fiscal2 ELSE NULL END) 			AS Fiscal2	,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Fiscal3 ELSE NULL END) 			AS Fiscal3	,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Fiscal4 ELSE NULL END) 			AS Fiscal4	,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Fiscal5 ELSE NULL END) 			AS Fiscal5	,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Des ELSE NULL END) 			AS Mon_Des	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Rec ELSE NULL END) 			AS Mon_Rec	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Por_Des ELSE NULL END) 			AS Por_Des	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Por_Rec ELSE NULL END) 			AS Por_Rec	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Otr1 ELSE NULL END) 			AS Mon_Otr1	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Otr2 ELSE NULL END) 			AS Mon_Otr2	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Mon_Otr3 ELSE NULL END) 			AS Mon_Otr3	, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN #tmpImpuestosFinal.Indice = 1 ")
            loComandoSeleccionar.AppendLine("				THEN #tmpLibroVentas.Referencia ELSE NULL END)			AS Referencia			 ")
            loComandoSeleccionar.AppendLine("FROM		#tmpLibroVentas ")
            loComandoSeleccionar.AppendLine("	JOIN #tmpImpuestosFinal")
            loComandoSeleccionar.AppendLine("		ON	#tmpImpuestosFinal.Documento = #tmpLibroVentas.Documento ")
            loComandoSeleccionar.AppendLine("		AND	#tmpImpuestosFinal.Cod_Tip = #tmpLibroVentas.Cod_Tip")
            loComandoSeleccionar.AppendLine("ORDER BY	Tabla ASC,")
            loComandoSeleccionar.AppendLine("			(CASE #tmpLibroVentas.Cod_Tip")
            loComandoSeleccionar.AppendLine("				WHEN 'FACT'		THEN 1")
            loComandoSeleccionar.AppendLine("				WHEN 'RETIVA'	THEN 2")
            loComandoSeleccionar.AppendLine("				WHEN 'N/CR'		THEN 3")
            loComandoSeleccionar.AppendLine("				WHEN 'N/DB'		THEN 4")
            loComandoSeleccionar.AppendLine("				ELSE				 5")
            loComandoSeleccionar.AppendLine("			END) ASC, #tmpLibroVentas.Fec_Ini ASC, ")
            loComandoSeleccionar.AppendLine("			#tmpImpuestosFinal.Documento, #tmpLibroVentas.Cod_Tip, #tmpImpuestosFinal.Indice")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Borra las tablas que quedan que no son necesarias")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpLibroVentas")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpImpuestosFinal")
            loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")


            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'-------------------------------------------------------------------
            ' Selección de opcion por excel (Microsoft Excel - xls)
			'-------------------------------------------------------------------
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Genera el archivo a partir de la tabla de datos y termina la ejecución
                Me.mGenerarArchivoExcel(laDatosReporte)

            End If


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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_USA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas_USA.ReportSource = loObjetoReporte

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

    Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

    '***********************************************************************'
    ' Prepara los datos para enviarlos al servicio web de Excel.            '
    '***********************************************************************'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)


    '***********************************************************************'
    ' Prepara los parámetros adicionales para enviarlos junto con los datos.'
    '***********************************************************************'
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje
        
        Dim loParametros As New NameValueCollection()
        loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
        loParametros.Add("lcRifEmpresa", cusAplicacion.goEmpresa.pcRifEmpresa)
        loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
        loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
        loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())

        Dim loClienteWeb As new WebClient()
        loClienteWeb.QueryString = loParametros
        
    '***********************************************************************'
    ' Envía los datos y parámetros, y espera la respuesta.                  '
    '***********************************************************************'
        Dim loRespuesta As Byte()  
        Try
            Dim lcRuta AS String = Me.MapPath("~\Framework\Xml\ParametrosGlobales.xml")
            DIm loParam As New System.Xml.XmlDocument()
            loParam.Load(lcRuta)
            Dim lcServicio As String = loParam.DocumentElement.GetAttribute("Servicios")

            loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rLibro_Ventas_USA_xlsx.aspx", loSalida.GetBuffer())
        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.ToString(), vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

    '***********************************************************************'
    ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel    '
    ' generado). Si el tipo está vacio : error desconocido.                 '
    '***********************************************************************'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type") 

        If String.IsNullOrEmpty(loTipoRespuesta) Then 
            'Error no especificado!
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
                                                                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then 

            Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta) 

            lcMensaje = Me.Server.HtmlEncode(lcMensaje)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        Else
            'Generación exitosa: la respuesta es el archivo en excel para descargar

            Me.Response.Clear()
            Me.Response.Buffer = True
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rMargen_gClienteVendedorArticulos_Resumido_MOL.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If


    End Sub


End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MAT: 07/10/11: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 22/11/11: Ajuste en signo en la totalización para las N/CR y Retenciones. Se			'
'				 quitaron los documentos de CxC anulados.									'
'-------------------------------------------------------------------------------------------'
' RJG: 05/10/12: Ajuste Totales: se colocaron en dos líneas para que no se monten los montos'
'-------------------------------------------------------------------------------------------'
' RJG: 04/09/14: Se adaptó para generar la salida personalizada a Excel por medio de un     '
'                servicio externo (eFactory Servicios).                                     '
'-------------------------------------------------------------------------------------------'
