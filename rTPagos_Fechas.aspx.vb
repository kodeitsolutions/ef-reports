'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTPagos_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rTPagos_Fechas
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Pagos.fec_ini) AS   Año,")
            loComandoSeleccionar.AppendLine("           DATEPART(MONTH, Pagos.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("           DATEPART(DAY, Pagos.fec_ini)AS Dia,")
            loComandoSeleccionar.AppendLine("           CASE")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Efectivo,")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Ticket,")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Cheque,")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Deposito,")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Detalles_Pagos.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("               ELSE 0")
            loComandoSeleccionar.AppendLine("           END AS Transferencia")
            loComandoSeleccionar.AppendLine(" INTO      #Temporal002 ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos Pagos")
            loComandoSeleccionar.AppendLine("           JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Pagos.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           JOIN Detalles_Pagos AS Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Pro  BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Status   IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Mon  BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Rev  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           GROUP BY DATEPART(YEAR, Pagos.fec_ini), DATEPART(MONTH, Pagos.fec_ini), DATEPART(DAY, Pagos.fec_ini), Detalles_Pagos.Tip_Ope")
            
            loComandoSeleccionar.AppendLine(" SELECT    'Cuantos'  AS  Cuantos, ")
            loComandoSeleccionar.AppendLine(" Año, ")
            loComandoSeleccionar.AppendLine(" Mes, ")
            loComandoSeleccionar.AppendLine(" Dia, ")
            loComandoSeleccionar.AppendLine(" SUM(Efectivo) As Efectivo, ")
            loComandoSeleccionar.AppendLine(" SUM(Ticket) As Ticket, ")
            loComandoSeleccionar.AppendLine(" SUM(Cheque) As Cheque, ")
            loComandoSeleccionar.AppendLine(" SUM(Tarjeta) As Tarjeta, ")
            loComandoSeleccionar.AppendLine(" SUM(Deposito) As Deposito, ")
            loComandoSeleccionar.AppendLine(" SUM(Transferencia) As Transferencia,")
            loComandoSeleccionar.AppendLine(" SUM(Efectivo) + SUM(Ticket) + SUM(Cheque) + SUM(Tarjeta) + SUM(Deposito) + SUM(Transferencia) AS Total_Pagos")
            loComandoSeleccionar.AppendLine(" FROM      #Temporal002 ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Año, Mes, Dia")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTPagos_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTPagos_Fechas.ReportSource = loObjetoReporte

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
' CMS: 22/06/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
