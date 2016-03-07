'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTConceptos_Movimiento_Mensual"
'-------------------------------------------------------------------------------------------'
Partial Class rTConceptos_Movimiento_Mensual
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	   
		    'Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia) 'Fecha
            'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1)) 'Semestre
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))'Concepto de Movimiento
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)) 'Sucursal
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3)) 
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) 'Revisión
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4)) 


            Dim lnAño As Integer = Cint(cusAplicacion.goReportes.paParametrosIniciales(0))
            If (lnAño = 0) Then
                lnAño = Date.Today.Year
            Else
                lnAño = Math.Min(Math.Max(lnAño, 1900), 2099)
            End If

            Dim llPrimerSemestre As Boolean = (CStr(cusAplicacion.goReportes.paParametrosIniciales(1)).Trim().ToLower() = "si")
            Dim lcPrimerSemestre As String = goServicios.mObtenerCampoFormatoSQL(llPrimerSemestre)
            
            Dim lcParametro0Desde As String 
            Dim lcParametro0Hasta As String 
            If llPrimerSemestre Then
                lcParametro0Desde = goServicios.mObtenerCampoFormatoSQL(New Date(lnAño, 1, 1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
                lcParametro0Hasta = goServicios.mObtenerCampoFormatoSQL(New Date(lnAño, 6, 30), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Else
                lcParametro0Desde = goServicios.mObtenerCampoFormatoSQL(New Date(lnAño, 7, 1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
                lcParametro0Hasta = goServicios.mObtenerCampoFormatoSQL(New Date(lnAño, 12, 31), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            End If

            Dim llOrdenesDePagoIncluyenImpuesto As Boolean = goOpciones.mObtener("INCIMPORDP","L")
            Dim lcOrdenesDePagoIncluyenImpuesto As String = goServicios.mObtenerCampoFormatoSQL(llOrdenesDePagoIncluyenImpuesto)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
	
			Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @llOrdenPago_IncluyeImpuesto BIT;")
            loConsulta.AppendLine("SET @llOrdenPago_IncluyeImpuesto = " & lcOrdenesDePagoIncluyenImpuesto & ";")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Conceptos.Cod_Con                                                       AS Cod_Con,")
            loConsulta.AppendLine("        Conceptos.Nom_Con                                                       AS Nom_Con,")
            loConsulta.AppendLine("        CAST(" & lcPrimerSemestre & " AS INT)                                   AS Semestre,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (1,7) THEN Monto ELSE 0 END)               AS Monto_01,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (2,8) THEN Monto ELSE 0 END)               AS Monto_02,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (3,9) THEN Monto ELSE 0 END)               AS Monto_03,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (4,10)THEN Monto ELSE 0 END)               AS Monto_04,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (5,11)THEN Monto ELSE 0 END)               AS Monto_05,")
            loConsulta.AppendLine("        SUM(CASE WHEN Totales.Mes IN (6,12)THEN Monto ELSE 0 END)               AS Monto_06,")
            loConsulta.AppendLine("        SUM(Monto)                                                              AS Monto_Total")
            loConsulta.AppendLine("FROM    (")
            loConsulta.AppendLine("            SELECT      CAST('Compra' AS VARCHAR(30))                           AS Origen,")
            loConsulta.AppendLine("                        Articulos.Cod_Con                                       AS Cod_Con,")
            loConsulta.AppendLine("                        MONTH(Compras.Fec_Ini)                                  AS Mes,")
            loConsulta.AppendLine("                        SUM(ROUND((")
            loConsulta.AppendLine("                            Renglones_Compras.Mon_Net")
            loConsulta.AppendLine("                            *(100-Compras.Por_Des1+Compras.Por_Rec1)/100")
            loConsulta.AppendLine("                            + Renglones_Compras.Mon_Imp1")
            loConsulta.AppendLine("                            + (CASE WHEN Compras.Mon_Bru>0 THEN")
            loConsulta.AppendLine("                                (   Compras.Mon_Otr1")
            loConsulta.AppendLine("                                  + Compras.Mon_Otr2")
            loConsulta.AppendLine("                                  + Compras.Mon_Otr3)")
            loConsulta.AppendLine("                              *Renglones_Compras.Mon_Net/Compras.Mon_Bru ")
            loConsulta.AppendLine("                              ELSE 0 END))*Compras.Tasa, 2))                    AS Monto")
            loConsulta.AppendLine("            FROM        Compras")
            loConsulta.AppendLine("                JOIN    Renglones_Compras ")
            loConsulta.AppendLine("                    ON  Renglones_Compras.Documento = Compras.Documento")
            loConsulta.AppendLine("                JOIN    Articulos ")
            loConsulta.AppendLine("                    ON  Articulos.Cod_Art = Renglones_Compras.Cod_Art")
            loConsulta.AppendLine("            WHERE       Compras.Status IN ('Confirmado', 'Afectado', 'Pagado')")
            loConsulta.AppendLine("                    AND Compras.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("                        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Rev BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("            GROUP BY    Articulos.Cod_Con, MONTH(Compras.Fec_Ini)")
            loConsulta.AppendLine("            UNION ALL")
            loConsulta.AppendLine("            SELECT      CAST('Orden de Pago' AS VARCHAR(30))                    AS Origen,")
            loConsulta.AppendLine("                        Renglones_oPagos.Cod_Con                                AS Cod_Con,")
            loConsulta.AppendLine("                        MONTH(Ordenes_Pagos.Fec_Ini)                            AS Mes,")
            loConsulta.AppendLine("                        SUM(ROUND((Renglones_oPagos.Mon_Net")
            loConsulta.AppendLine("                        + (CASE WHEN @llOrdenPago_IncluyeImpuesto = 0 ")
            loConsulta.AppendLine("                            THEN Renglones_oPagos.Mon_Imp1")
            loConsulta.AppendLine("                            ELSE 0 END))*Ordenes_Pagos.Tasa, 2))                AS Monto")
            loConsulta.AppendLine("            FROM        Ordenes_Pagos")
            loConsulta.AppendLine("                JOIN    Renglones_oPagos ")
            loConsulta.AppendLine("                    ON  Renglones_oPagos.Documento = Ordenes_Pagos.Documento")
            loConsulta.AppendLine("            WHERE       Ordenes_Pagos.Status IN ('Confirmado')")
            loConsulta.AppendLine("                    AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                    AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("                        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                    AND Ordenes_Pagos.Cod_Rev BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("            GROUP BY    Renglones_oPagos.Cod_Con, MONTH(Ordenes_Pagos.Fec_Ini)")
            loConsulta.AppendLine("            UNION ALL")
            loConsulta.AppendLine("            SELECT      CAST('Movimiento de Cuenta' AS VARCHAR(30))             AS Origen,")
            loConsulta.AppendLine("                        Movimientos_Cuentas.Cod_Con                             AS Cod_Con,")
            loConsulta.AppendLine("                        MONTH(Movimientos_Cuentas.Fec_Ini)                      AS Mes,")
            loConsulta.AppendLine("                        SUM(ROUND((Movimientos_Cuentas.Mon_Deb ")
            loConsulta.AppendLine("                        - Movimientos_Cuentas.Mon_Hab)")
            loConsulta.AppendLine("                            *Movimientos_Cuentas.Tasa, 2))                      AS Monto")
            loConsulta.AppendLine("            FROM        Movimientos_Cuentas")
            loConsulta.AppendLine("            WHERE       Movimientos_Cuentas.Status IN ('Confirmado')")
            loConsulta.AppendLine("                    AND Movimientos_Cuentas.Automatico = 0")
            loConsulta.AppendLine("                    AND Movimientos_Cuentas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                    AND Movimientos_Cuentas.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("                        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                    AND Movimientos_Cuentas.Cod_Rev BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("            GROUP BY    Movimientos_Cuentas.Cod_Con, MONTH(Movimientos_Cuentas.Fec_Ini)")
            loConsulta.AppendLine("            UNION ALL")
            loConsulta.AppendLine("            SELECT      CAST('Movimiento de Caja' AS VARCHAR(30))               AS Origen,")
            loConsulta.AppendLine("                        Movimientos_Cajas.Cod_Con                               AS Cod_Con,")
            loConsulta.AppendLine("                        MONTH(Movimientos_Cajas.Fec_Ini)                        AS Mes,")
            loConsulta.AppendLine("                        SUM(ROUND((Movimientos_Cajas.Mon_Deb ")
            loConsulta.AppendLine("                        - Movimientos_Cajas.Mon_Hab)")
            loConsulta.AppendLine("                            *Movimientos_Cajas.Tasa, 2))                        AS Monto")
            loConsulta.AppendLine("            FROM        Movimientos_Cajas")
            loConsulta.AppendLine("            WHERE       Movimientos_Cajas.Status IN ('Confirmado')")
            loConsulta.AppendLine("                    AND Movimientos_Cajas.Automatico = 0")
            loConsulta.AppendLine("                    AND Movimientos_Cajas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                    AND Movimientos_Cajas.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("                        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                    AND Movimientos_Cajas.Cod_Rev BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("            GROUP BY    Movimientos_Cajas.Cod_Con, MONTH(Movimientos_Cajas.Fec_Ini)")
            loConsulta.AppendLine(") AS Totales")
            loConsulta.AppendLine("    JOIN Conceptos ON Conceptos.Cod_Con = Totales.Cod_Con")
            loConsulta.AppendLine("WHERE    Conceptos.Cod_Con BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("GROUP BY Conceptos.Cod_Con, Conceptos.Nom_Con")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTConceptos_Movimiento_Mensual", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTConceptos_Movimiento_Mensual.ReportSource =	 loObjetoReporte	

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
' RJG: 18/06/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
