'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCobros_Año"
'-------------------------------------------------------------------------------------------'
Partial Class rTCobros_Año
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()





            loComandoSeleccionar.AppendLine("   SELECT ")
            loComandoSeleccionar.AppendLine("               datepart(YEAR, fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(Mon_Net,0)) AS Mon_Net")
            loComandoSeleccionar.AppendLine("   INTO #Temporal")
            loComandoSeleccionar.AppendLine("   FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("   WHERE   Cod_Tip = 'Fact'")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro6Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            
            loComandoSeleccionar.AppendLine("   GROUP BY DATEPART(YEAR, fec_ini)")

            loComandoSeleccionar.AppendLine("   SELECT")
            loComandoSeleccionar.AppendLine("               datepart(YEAR, Cobros.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Efectivo,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Ticket,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Cheque,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Deposito,")
            loComandoSeleccionar.AppendLine("               CASE")
            loComandoSeleccionar.AppendLine("                   WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("                   ELSE 0")
            loComandoSeleccionar.AppendLine("               END AS Transferencia")
            loComandoSeleccionar.AppendLine("   INTO #Temporal3 ")
            loComandoSeleccionar.AppendLine("   FROM Cobros Cobros")
            loComandoSeleccionar.AppendLine("   JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine("   JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("   WHERE   Cobros.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cobros.Cod_Rev between " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cobros.Cod_Rev NOT between " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            
            loComandoSeleccionar.AppendLine("           GROUP BY DATEPART(YEAR, Cobros.fec_ini), Detalles_Cobros.Tip_Ope")

            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("   		Año,  ")
            loComandoSeleccionar.AppendLine("   		SUM(Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("   		SUM(Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("   		SUM(Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("   		SUM(Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("   		SUM(Deposito) AS Deposito, ")
            loComandoSeleccionar.AppendLine("   		SUM(Transferencia) AS Transferencia ")
            loComandoSeleccionar.AppendLine("   INTO #Temporal2  ")
            loComandoSeleccionar.AppendLine("   FROM #Temporal3  ")
            loComandoSeleccionar.AppendLine("   GROUP BY Año")

            loComandoSeleccionar.AppendLine("   SELECT    ")
            loComandoSeleccionar.AppendLine("               ISNULL(#temporal.Año, #temporal2.Año) AS Año,    ")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal.Mon_Net,0)) AS Mon_Net,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Efectivo,0)) AS Efectivo,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Ticket,0)) AS Ticket,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Cheque,0))   AS Cheque,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Tarjeta,0))  AS Tarjeta,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Deposito,0)) AS Deposito,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Transferencia,0)) AS Transferencia,")
            loComandoSeleccionar.AppendLine("               SUM(ISNULL(#temporal2.Efectivo,0) + ISNULL(#temporal2.Cheque,0) + ISNULL(#temporal2.Tarjeta,0) + ISNULL(#temporal2.Deposito,0) + ISNULL(#temporal2.Transferencia,0) + ISNULL(#temporal2.Ticket,0)) AS Total_Cobros")
            loComandoSeleccionar.AppendLine("   FROM	#Temporal #Temporal   ")
            loComandoSeleccionar.AppendLine("   FULL JOIN  #temporal2 AS #temporal2 ON ((#temporal.Año = #temporal2.Año) ) ")
            loComandoSeleccionar.AppendLine("   GROUP BY ISNULL(#temporal.Año, #temporal2.Año)   ")
            loComandoSeleccionar.AppendLine("   ORDER BY " & lcOrdenamiento )


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCobros_Año", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCobros_Año.ReportSource = loObjetoReporte

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
' CMS: 23/06/10: Programacion inicial
'-------------------------------------------------------------------------------------------'