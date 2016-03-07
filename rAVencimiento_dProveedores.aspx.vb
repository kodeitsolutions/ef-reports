'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAVencimiento_dProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rAVencimiento_dProveedores
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            'Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Hasta As String = cusAplicacion.goReportes.paParametrosFinales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosFinales(12)
            Dim lcParametro13Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))

            Dim Fecha As String

            If lcParametro10Hasta = "Vencimiento" Then
                Fecha = "Fec_Fin"
            Else
                Fecha = "Fec_Ini"
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("  SELECT   ")
            loComandoSeleccionar.AppendLine("  ROW_NUMBER() OVER(PARTITION BY Proveedores.Cod_Pro ORDER BY Proveedores.Cod_Pro, Cuentas_Pagar.Documento, Cuentas_Pagar.Fec_Ini, Cuentas_Pagar.Fec_Fin) AS Fila, ")
            loComandoSeleccionar.AppendLine("  		Proveedores.Cod_Pro,   ")
            loComandoSeleccionar.AppendLine("  		Proveedores.Nom_Pro,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN Cuentas_Pagar.Mon_Sal > 0 THEN 1   ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS Cant_Doc,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN DATEDIFF(d,Cuentas_Pagar." & Fecha & " , " & lcParametro0Hasta & ") < 1     THEN ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Pagar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS Por_Vencer,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") >= 1) AND (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") <= 30) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Pagar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS a,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") >= 31) AND (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") <= 60) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Pagar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS b,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") >= 61) AND (DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") <= 90) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Pagar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS c,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") >= 91 THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Pagar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS D,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.Cod_Tip,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.Documento,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.Fec_Ini,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.control,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.factura,   ")
            loComandoSeleccionar.AppendLine("  		Cuentas_Pagar.Fec_Fin,   ")
            loComandoSeleccionar.AppendLine("  		DATEDIFF(d, Cuentas_Pagar." & Fecha & ", " & lcParametro0Hasta & ") AS Dias,   ")
            loComandoSeleccionar.AppendLine("  		(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Net *(-1) Else Cuentas_Pagar.Mon_Net End) As Mon_Net, ")
            loComandoSeleccionar.AppendLine("  		CASE    ")
            loComandoSeleccionar.AppendLine("  		    WHEN (DATALENGTH(Cuentas_Pagar.comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) THEN '- '+CAST(Cuentas_Pagar.comentario AS  VARCHAR(1000))+CHAR(13)+'- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("  		    WHEN (DATALENGTH(Cuentas_Pagar.comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) <= 1) THEN '- '+CAST(Cuentas_Pagar.comentario AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("  		    WHEN (DATALENGTH(Cuentas_Pagar.comentario) <= 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) THEN '- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("  		    ELSE ''    ")
            loComandoSeleccionar.AppendLine("  		END AS Comentario,    ")
            loComandoSeleccionar.AppendLine("  		Proveedores.Telefonos,   ")
            loComandoSeleccionar.AppendLine("  		Proveedores.Fax   ")
            loComandoSeleccionar.AppendLine("  FROM	Cuentas_Pagar   ")
            loComandoSeleccionar.AppendLine("  JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro   ")
            loComandoSeleccionar.AppendLine(" WHERE         Cuentas_Pagar.Mon_Sal <> 0 ")
            loComandoSeleccionar.AppendLine("               AND Cuentas_Pagar.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("      			AND Cuentas_Pagar.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Pro between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Ven between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proveedores.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Mon between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proveedores.Cod_Tip between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proveedores.Cod_Cla between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND             ((" & lcParametro7Desde & " = 'Si' AND (DATEDIFF(d, Cuentas_Pagar.Fec_Fin, " & lcParametro0Hasta & ") >= 1))")
            loComandoSeleccionar.AppendLine("               OR              (" & lcParametro7Desde & " <> 'Si' AND ((DATEDIFF(d, Cuentas_Pagar.Fec_Fin, " & lcParametro0Hasta & ") >= 1) or (DATEDIFF(d, Cuentas_Pagar.Fec_Fin, " & lcParametro0Hasta & ") < 1))))")
            loComandoSeleccionar.AppendLine("               AND Cuentas_Pagar.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("               AND " & lcParametro8Hasta)

            If lcParametro12Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev between " & lcParametro11Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT between " & lcParametro11Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("				AND Proveedores.Status In (" & lcParametro13Desde & ")")

            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento & ", Cuentas_Pagar.Documento, Cuentas_Pagar.Fec_Ini, Cuentas_Pagar.Fec_Fin ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAVencimiento_dProveedores", laDatosReporte)

            If lcParametro9Desde.ToString = "Si" Then
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line5"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line5"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            Else
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text2").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text3").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text2").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text3").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
                loObjetoReporte.ReportDefinition.Sections(3).Height = 300
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrAVencimiento_dProveedores.ReportSource = loObjetoReporte

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
' CMS: 30/06/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  16/07/09: Filtro Calculado por, verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  18/07/09: Se agrego la columna Dias
'-------------------------------------------------------------------------------------------'
' CMS:  31/07/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  23/06/10: Filtro “Estatus del Proveedor:” 
'-------------------------------------------------------------------------------------------'