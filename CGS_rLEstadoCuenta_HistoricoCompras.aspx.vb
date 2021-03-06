﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLEstadoCuenta_HistoricoCompras"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLEstadoCuenta_HistoricoCompras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))


            Dim Empresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10)	= 0")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio AS VARCHAR(10) = ''")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Saldo Inicial")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE -Cuentas_Pagar.Mon_Net	")
            loComandoSeleccionar.AppendLine("			END) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("INTO	#tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            If Empresa = "'Cegasa'" Then
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip = 'FACT'")
            Else
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip IN ('FACT','N/CR')")
            End If
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Reg < @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND (Cuentas_Pagar.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE -Cuentas_Pagar.Mon_Net	")
            loComandoSeleccionar.AppendLine("			END) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            If Empresa = "'Cegasa'" Then
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip <> 'FACT'")
            Else
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip NOT IN ('FACT','N/CR')")
            End If
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini < @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND (Cuentas_Pagar.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN -Renglones_Pagos.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Pagos.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END) +(Pagos.Mon_Ret + Pagos.Mon_Des) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	Pagos")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Pagos.Fec_Ini < @ldFecha_Desde")
            'loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Pro BETWEEN @sp_CodPro_Desde AND @sp_CodPro_Hasta")
            loComandoSeleccionar.AppendLine("		AND Pagos.Automatico = 0")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Pagos.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("		AND (Pagos.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro,Pagos.Mon_Ret,Pagos.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Pagar'									AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura							AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg							AS Registro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control							AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("				THEN @lnCero")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            If Empresa = "'Cegasa'" Then
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip = 'FACT'")
            Else
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip IN ('FACT','N/CR')")
            End If
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND (Cuentas_Pagar.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Pagar'									AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Registro,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Cod_Tip = 'ADEL' THEN")
            loComandoSeleccionar.AppendLine("            (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Pagos")
            loComandoSeleccionar.AppendLine("            WHERE Documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            ) ")
            loComandoSeleccionar.AppendLine("            ELSE SUBSTRING(Cuentas_Pagar.Comentario,0,50)")
            loComandoSeleccionar.AppendLine("        END											AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("				THEN @lnCero")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            If Empresa = "'Cegasa'" Then
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip <> 'FACT'")
            Else
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip NOT IN ('FACT','N/CR')")
            End If
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND (Cuentas_Pagar.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cuentas_Pagar.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Pagos'											AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		'PAGO'	    									AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Pagos.Documento									AS Documento,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini									AS Registro,")
            loComandoSeleccionar.AppendLine("		CASE WHEN (SELECT COUNT(Documento) FROM Detalles_Pagos AS D_P WHERE D_P.Documento = Pagos.Documento) > 1")
            loComandoSeleccionar.AppendLine("			THEN SUBSTRING(Pagos.Comentario,0,35)")
            loComandoSeleccionar.AppendLine("			ELSE (CONCAT((SELECT CONCAT(RTRIM(D_P.Tip_Ope),' ',RTRIM(D_P.Num_Doc)) ")
            loComandoSeleccionar.AppendLine("							FROM Detalles_Pagos AS D_P WHERE D_P.Documento = Pagos.Documento),")
            loComandoSeleccionar.AppendLine("				  ' ', '/ ', SUBSTRING(Pagos.Comentario, 6, 30)) )")
            loComandoSeleccionar.AppendLine("		END										        AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	Renglones_Pagos.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE	@lnCero	")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	@lnCero")
            loComandoSeleccionar.AppendLine("				ELSE	Renglones_Pagos.Mon_Abo	")
            If Empresa = "'Cegasa'" Then
                loComandoSeleccionar.AppendLine("			END) +(Pagos.Mon_Ret + Pagos.Mon_Des) 		AS Mon_Hab,")
            Else
                loComandoSeleccionar.AppendLine("			END) + Pagos.Mon_Ret - Pagos.Mon_Des 		AS Mon_Hab,")
            End If
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Pagos")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Pagos.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("		AND	Pagos.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Pagos.Automatico = 0")
            loComandoSeleccionar.AppendLine("		AND (Pagos.Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Pagos.Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("GROUP BY	Proveedores.Cod_Pro, Proveedores.Nom_Pro, Pagos.Cod_Ven, Pagos.Documento,")
            loComandoSeleccionar.AppendLine("			 Pagos.Fec_Ini, Pagos.Registro, Pagos.Comentario, Pagos.Mon_Ret, Pagos.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Pendientes por conciliar")
            loComandoSeleccionar.AppendLine("SELECT COUNT(Documento)									AS Pendientes,")
            loComandoSeleccionar.AppendLine("		Cod_Pro											AS Cod_Pro")
            loComandoSeleccionar.AppendLine("INTO #tmpMovPendientesFact")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'FACT' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loComandoSeleccionar.AppendLine("		AND (Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("	AND Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT COUNT(Documento)								AS Pendientes,")
            loComandoSeleccionar.AppendLine("		Cod_Pro											AS Cod_Pro")
            loComandoSeleccionar.AppendLine("INTO #tmpMovPendientesAdel")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'ADEL' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loComandoSeleccionar.AppendLine("		AND (Cod_Pro = " & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("		OR Cod_Pro = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("	AND Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("SET		Orden = M.Orden,")
            loComandoSeleccionar.AppendLine("		Mon_Sal = -M.Mon_Deb+M.Mon_Hab,")
            loComandoSeleccionar.AppendLine("		Sal_Ini = M.Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	ROW_NUMBER() ")
            loComandoSeleccionar.AppendLine("						OVER (	PARTITION BY #tmpMovimientos.Cod_Pro ")
            loComandoSeleccionar.AppendLine("								ORDER BY #tmpMovimientos.Fec_Ini, (CASE WHEN #tmpMovimientos.Cod_Tip='' THEN 'zzzzzzzzz' ELSE #tmpMovimientos.Cod_Tip END ) ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Tabla, #tmpMovimientos.Cod_Tip, #tmpMovimientos.Documento,")
            loComandoSeleccionar.AppendLine("					ISNULL(SI.Sal_Ini, @lnCero) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Deb AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Hab AS Mon_Hab")
            loComandoSeleccionar.AppendLine("			FROM	#tmpMovimientos			")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Pro, SUM(Sal_Ini) AS Sal_Ini FROM #tmpSaldos_Iniciales GROUP BY Cod_Pro) AS SI")
            loComandoSeleccionar.AppendLine("				ON SI.Cod_Pro = #tmpMovimientos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		) AS M		")
            loComandoSeleccionar.AppendLine("WHERE	M.Tabla = #tmpMovimientos.Tabla ")
            loComandoSeleccionar.AppendLine("	AND	M.Cod_Tip = #tmpMovimientos.Cod_Tip")
            loComandoSeleccionar.AppendLine("	AND	M.Documento = #tmpMovimientos.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Tabla, A.Cod_Pro, A.Nom_Pro, A.Cod_Tip, A.Documento,")
            loComandoSeleccionar.AppendLine("		A.Fec_Ini, A.Registro, A.Sal_Ini, A.Mon_Deb, A.Mon_Hab,")
            loComandoSeleccionar.AppendLine("       CASE WHEN LEN(A.Referencia) > 50")
            loComandoSeleccionar.AppendLine("            THEN CONCAT(SUBSTRING(A.Referencia,1,50),'...')")
            loComandoSeleccionar.AppendLine("            ELSE A.Referencia")
            loComandoSeleccionar.AppendLine("       END     AS Referencia,")
            loComandoSeleccionar.AppendLine("		SUM(B.Mon_Sal) +  A.Sal_Ini AS Sal_Doc, @ldFecha_Desde AS Desde, @ldFecha_Hasta	AS Hasta,")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesFact.Pendientes, @lnCero) AS PendientesFact,")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesAdel.Pendientes, @lnCero) AS PendientesAdel")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B")
            loComandoSeleccionar.AppendLine("		ON B.Cod_Pro = A.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesFact ON #tmpMovPendientesFact.Cod_Pro = A.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesAdel ON #tmpMovPendientesAdel.Cod_Pro = A.Cod_Pro")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Tabla, A.Cod_Pro, A.Nom_Pro, A.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		A.Documento, A.Fec_Ini, A.Registro, A.Referencia, A.Sal_Ini, ")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab, #tmpMovPendientesFact.Pendientes, #tmpMovPendientesAdel.Pendientes")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  0 AS Orden, 'Saldo Inicial' AS Tabla, #tmpSaldos_Iniciales.Cod_Pro, Proveedores.Nom_Pro AS Nom_Pro, 'NR' AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		@lcVacio AS Documento, @lcVacio AS Fec_Ini, @lcVacio AS Registro, SUM(Sal_Ini) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		@lnCero AS Mon_Deb, @lnCero AS Mon_Hab, 'SIN MOVIMIENTOS' AS Referencia, @lnCero AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("		@ldFecha_Desde AS Desde, @ldFecha_Hasta	AS Hasta, ")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesFact.Pendientes, @lnCero) AS PendientesFact,")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesAdel.Pendientes, @lnCero) AS PendientesAdel")
            loComandoSeleccionar.AppendLine("FROM #tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("  JOIN Proveedores ON #tmpSaldos_Iniciales.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesFact ON #tmpMovPendientesFact.Cod_Pro = #tmpSaldos_Iniciales.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesAdel ON #tmpMovPendientesAdel.Cod_Pro = #tmpSaldos_Iniciales.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE #tmpSaldos_Iniciales.Cod_Pro NOT IN (SELECT Cod_Pro FROM #tmpMovimientos)")
            loComandoSeleccionar.AppendLine("GROUP BY #tmpSaldos_Iniciales.Cod_Pro, Proveedores.Nom_Pro, #tmpMovPendientesFact.Pendientes, #tmpMovPendientesAdel.Pendientes")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro ASC, Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovPendientesFact")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovPendientesAdel")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLEstadoCuenta_HistoricoCompras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rLEstadoCuenta_HistoricoCompras.ReportSource = loObjetoReporte


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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
