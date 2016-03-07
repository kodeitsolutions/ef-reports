'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAVencimiento_tVendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rAVencimiento_tVendedores
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
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Hasta As String = cusAplicacion.goReportes.paParametrosFinales(9)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosFinales(11)
            Dim lcParametro12Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))

            Dim Fecha As String

            If lcParametro9Hasta = "Vencimiento" Then
                Fecha = "Fec_Fin"
            Else
                Fecha = "Fec_Ini"
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("  SELECT   ")
            loComandoSeleccionar.AppendLine("  		Vendedores.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("  		Vendedores.Nom_Ven,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN Cuentas_Cobrar.Mon_Sal > 0 THEN 1   ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS Cant_Doc,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN DATEDIFF(d, Cuentas_Cobrar." & Fecha & " , " & lcParametro0Hasta & ") < 1 THEN ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS Por_Vencer,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") >= 1) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") <= 30) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS a,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") >= 31) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") <= 60) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS b,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") >= 61) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") <= 90) THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS c,   ")
            loComandoSeleccionar.AppendLine("  		CASE   ")
            loComandoSeleccionar.AppendLine("  			WHEN DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro0Hasta & ") >= 91 THEN    ")
            loComandoSeleccionar.AppendLine("  		        CASE    ")
            loComandoSeleccionar.AppendLine("  		        		WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal*(-1)    ")
            loComandoSeleccionar.AppendLine("  		        		ELSE Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("  		        End    ")
            loComandoSeleccionar.AppendLine("  			ELSE 0   ")
            loComandoSeleccionar.AppendLine("  		END AS d   ")
            loComandoSeleccionar.AppendLine("  INTO	#tEMPORAL   ")
            loComandoSeleccionar.AppendLine("  FROM	Cuentas_Cobrar   ")
            loComandoSeleccionar.AppendLine("  JOIN Vendedores ON Vendedores.Cod_Ven = Cuentas_Cobrar.Cod_Ven   ")
            loComandoSeleccionar.AppendLine(" WHERE         cuentas_cobrar.Mon_Sal <> 0 ")
            loComandoSeleccionar.AppendLine("               AND cuentas_cobrar.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("      			AND cuentas_cobrar.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND cuentas_cobrar.Cod_Ven between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND cuentas_cobrar.Cod_Ven between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND cuentas_cobrar.Cod_Mon between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Tip between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Vendedores.Clase between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND             ((" & lcParametro7Desde & " = 'Si' AND (DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro0Hasta & ") >= 1))")
            loComandoSeleccionar.AppendLine("               OR              (" & lcParametro7Desde & " <> 'Si' AND ((DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro0Hasta & ") >= 1) or (DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro0Hasta & ") < 1))))")
            loComandoSeleccionar.AppendLine("               AND	Cuentas_Cobrar.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("               AND	" & lcParametro8Hasta)

            If lcParametro11Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro10Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro10Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("     AND Vendedores.Status In (" & lcParametro12Desde & ")")


            loComandoSeleccionar.AppendLine("  SELECT   ")
            loComandoSeleccionar.AppendLine("  			#Temporal.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("  			#Temporal.Nom_Ven,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.Cant_Doc) AS Cant_Doc,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.Por_Vencer) AS Por_Vencer,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.a) AS A,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.b) AS B,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.c) AS C,   ")
            loComandoSeleccionar.AppendLine("  			SUM(#Temporal.d) AS D   ")
            loComandoSeleccionar.AppendLine("  FROM #Temporal   ")
            loComandoSeleccionar.AppendLine("  GROUP BY #Temporal.Cod_Ven, #Temporal.Nom_Ven   ")
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)



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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAVencimiento_tVendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrAVencimiento_tVendedores.ReportSource = loObjetoReporte

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
' CMS: 09/06/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  16/07/09: Filtro Calculado por, verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  31/07/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  23/06/10: Filtro “Estatus del Vendedor:”
'-------------------------------------------------------------------------------------------'