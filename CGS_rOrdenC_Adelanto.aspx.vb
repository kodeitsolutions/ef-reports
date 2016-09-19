'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOrdenC_Adelanto"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOrdenC_Adelanto
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
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde	DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta	DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProDesde	VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProHasta	VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcOrdenDesde	VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcOrdenHasta	VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Ordenes_Compras.Documento	AS Documento,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini		AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net		AS Monto,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Sal		AS Saldo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario	AS Comentario,")
            loComandoSeleccionar.AppendLine("		Pagos.Documento				AS Pago,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini				AS Fecha_Pago,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento		AS Adelanto,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini		AS Fecha_Adel,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Mon_Net		AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net - SUM(Cuentas_Pagar.Mon_Net) AS Deuda")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Pagos ON Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip = 'ADEL'")
            loComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini > '01/06/2016'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Status <> 'Pagado'")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento BETWEEN @lcOrdenDesde AND @lcOrdenHasta")
            loComandoSeleccionar.AppendLine("GROUP BY Ordenes_Compras.Documento, Ordenes_Compras.Fec_Ini, Ordenes_Compras.Cod_Pro, Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net, Ordenes_Compras.Mon_Sal, Ordenes_Compras.Comentario, Pagos.Documento,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini, Cuentas_Pagar.Documento, Cuentas_Pagar.Fec_Ini, Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("ORDER BY Ordenes_Compras.Documento")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOrdenC_Adelanto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOrdenC_Adelanto.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

