'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumen_StatusDocumentos"
'-------------------------------------------------------------------------------------------'
Partial Class rResumen_StatusDocumentos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosIniciales(3)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Ajustes de Inventario' As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ajustes")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("		AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine("		AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("		And " & lcParametro2Hasta)	            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Traslados entre Almacenes'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Traslados")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Ajustes de Precios'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ajustes_precios")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Cortes de Inventarios'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Cortes")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)	            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Formas Libres de Inventario'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM libres_Inventarios")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Facturas de Ventas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Facturas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Cotizaciones'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Cotizaciones")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Pedidos'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Pedidos")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Notas de Entregas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Entregas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Guias de Despacho'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Guias")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Devoluciones de Clientes'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM devoluciones_clientes")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)	              
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Cuentas por Cobrar'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Cuentas_Cobrar")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Cobros'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Cobros")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Formas Libres de Ventas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Libres_Ventas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Presupuestos'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Presupuestos")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Ordenes de Compras'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ordenes_Compras")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Notas de Recepcion'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Recepciones")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Requisiciones Internas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Requisiciones")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Facturas de Compras'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Compras")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Devoluciones a Proveedores'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Devoluciones_Proveedores")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Cuentas por Pagar'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Cuentas_Pagar")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Pagos a Proveedores'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Pagos")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)	             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Formas Libres de Ventas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Libres_Ventas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Ordenes de Pagos'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ordenes_Pagos")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmao, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Movimientos de Cuentas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Movimientos_Cuentas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Movimientos de Cajas'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Movimientos_Cajas")
			loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)	             
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Pendiente' THEN 1 ELSE 0 END),0) AS  Pendiente, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Anulado' THEN 1 ELSE 0 END),0) AS  Anulado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Confirmado' THEN 1 ELSE 0 END),0) AS  Confirmado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Procesado' THEN 1 ELSE 0 END),0) AS  Procesado, ")
			loComandoSeleccionar.AppendLine(" 	ISNULL(SUM(CASE WHEN Status = 'Afectado' THEN 1 ELSE 0 END),0) AS  Afectado, ")
			loComandoSeleccionar.AppendLine(" 	'Depositos Bancarios'  As Tabla")
			loComandoSeleccionar.AppendLine(" FROM Depositos")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro2Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.AppendLine("           AND Cod_Mon Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumen_StatusDocumentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrResumen_StatusDocumentos.ReportSource = loObjetoReporte

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
' CMS: 06/10/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'