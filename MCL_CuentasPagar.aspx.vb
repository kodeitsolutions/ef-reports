'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_CuentasPagar"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_CuentasPagar
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim ldDecimales As Decimal = goOpciones.pnDecimalesParaMonto

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine("DECLARE @ldFecFact_Desde AS DATETIME = " & lcParametro0Desde)
            'loComandoSeleccionar.AppendLine("DECLARE @ldFecFact_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecFact_Hasta AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcMes AS VARCHAR(15) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoPAT AS DECIMAL(28," & ldDecimales & ") =  " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoISLR AS DECIMAL(28," & ldDecimales & ") = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoFAOV AS DECIMAL(28," & ldDecimales & ") = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoIVSS AS DECIMAL(28," & ldDecimales & ") = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoSEG AS DECIMAL(28," & ldDecimales & ") =  " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldMontoOTRO AS DECIMAL(28," & ldDecimales & ") = " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcDescOtro AS VARCHAR(MAX) = " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("Select	Facturas.Fec_Ini				AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Facturas.Factura				AS Factura,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro				AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Facturas.Mon_Bas1				AS Mon_Bas,")
            loComandoSeleccionar.AppendLine("		Facturas.Mon_Imp1				AS Mon_Imp,")
            loComandoSeleccionar.AppendLine("		Facturas.Mon_Net				AS Mon_Net,")
            loComandoSeleccionar.AppendLine("		Facturas.Mon_Sal				AS Mon_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(ISLR.Mon_Net,0)		AS MonNet_ISLR,")
            loComandoSeleccionar.AppendLine("		COALESCE(Patente.Mon_Net,0)		AS MonNet_PAT,")
            loComandoSeleccionar.AppendLine("		RTRIM(@lcMes) 		            AS Mes_Montos,")
            loComandoSeleccionar.AppendLine("		@ldMontoPAT 		            AS MontoPAT,")
            loComandoSeleccionar.AppendLine("		@ldMontoISLR		            AS MontoISLR,")
            loComandoSeleccionar.AppendLine("		@ldMontoFAOV		            AS MontoFAOV,")
            loComandoSeleccionar.AppendLine("		@ldMontoIVSS		            AS MontoIVSS,")
            loComandoSeleccionar.AppendLine("		@ldMontoSEG 		            AS MontoSEG,")
            loComandoSeleccionar.AppendLine("		@ldMontoOTRO		            AS MontoOTRO,")
            loComandoSeleccionar.AppendLine("		@lcDescOtro		                AS DescOtro")
            'loComandoSeleccionar.AppendLine("		COALESCE(SUM(Adelantos.Mon_Net),0)	As MonNet_ADEL")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar As Facturas")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores On Proveedores.Cod_Pro = Facturas.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar As ISLR On Facturas.documento = ISLR.Doc_Ori")
            loComandoSeleccionar.AppendLine("		And Facturas.Cod_Tip = ISLR.Cla_Ori")
            loComandoSeleccionar.AppendLine("		And ISLR.Cod_Tip = 'ISLR' AND ISLR.Status <> 'Pagado'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar AS Patente ON Facturas.documento = Patente.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Tip = Patente.Cla_Ori")
            loComandoSeleccionar.AppendLine("		AND Patente.Cod_Tip = 'RETPAT'AND Patente.Status <> 'Pagado'")
            'loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar AS Adelantos ON Proveedores.Cod_Pro = Adelantos.Cod_Pro")
            'loComandoSeleccionar.AppendLine("		AND Adelantos.Status <> 'Pagado'")
            'loComandoSeleccionar.AppendLine("		AND Adelantos.Fec_Ini < @ldFecFact_Hasta")
            loComandoSeleccionar.AppendLine("WHERE Facturas.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Status <> 'Pagado'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Mon_Sal > 0")
            loComandoSeleccionar.AppendLine("	AND Facturas. Fec_Reg < @ldFecFact_Hasta")
            'loComandoSeleccionar.AppendLine("GROUP BY Facturas.Fec_Ini,Facturas.Factura,Proveedores.Nom_Pro,Facturas.Mon_Bas1,Facturas.Mon_Imp1,Facturas.Mon_Net,")
            'loComandoSeleccionar.AppendLine("		 ISLR.Mon_Net,Patente.Mon_Net")
            loComandoSeleccionar.AppendLine("ORDER BY Facturas.Fec_Ini")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_CuentasPagar", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvMCL_CuentasPagar.ReportSource = loObjetoReporte

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
' KDE: 14/09/17: Codigo inicial
'-------------------------------------------------------------------------------------------'
