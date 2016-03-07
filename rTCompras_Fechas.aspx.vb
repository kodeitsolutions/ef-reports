'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCompras_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rTCompras_Fechas
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("WITH curTemporal AS ( ")

            loComandoSeleccionar.AppendLine("SELECT	CONVERT(nchar(10), Compras.Fec_Ini, 103)	AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       CONVERT(nchar(10), Compras.Fec_Ini, 112)	AS	Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("       Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("       Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("       Compras.Mon_Imp2, ")
            loComandoSeleccionar.AppendLine("       Compras.Mon_Imp3, ")
            loComandoSeleccionar.AppendLine("       'Comprado'       AS	Tipo")
            loComandoSeleccionar.AppendLine("FROM	Compras ")
            loComandoSeleccionar.AppendLine("WHERE	Compras.Status IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("		AND 	Compras.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND		Compras.Cod_Pro      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND		Compras.Cod_Ven      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND		Compras.Cod_Rev      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND		Compras.Cod_Suc      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro4Hasta)


            ' Union con Select de la tabla de Devoluciones
            loComandoSeleccionar.AppendLine(" UNION ALL ")

            loComandoSeleccionar.AppendLine("SELECT			CONVERT(nchar(10), Devoluciones_Proveedores.Fec_Ini, 103)	AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("				CONVERT(nchar(10), Devoluciones_Proveedores.Fec_Ini, 112)	AS	Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("				Devoluciones_Proveedores.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("				Devoluciones_Proveedores.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("				Devoluciones_Proveedores.Mon_Imp2, ")
            loComandoSeleccionar.AppendLine("				Devoluciones_Proveedores.Mon_Imp3, ")
            loComandoSeleccionar.AppendLine("				'Devuelto '       AS	Tipo")
            loComandoSeleccionar.AppendLine("FROM			Devoluciones_Proveedores ")
            loComandoSeleccionar.AppendLine("WHERE			Devoluciones_Proveedores.Status       IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("       AND		Devoluciones_Proveedores.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND		Devoluciones_Proveedores.Cod_Pro      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		AND	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND		Devoluciones_Proveedores.Cod_Ven      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		AND	" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" ) ")

            loComandoSeleccionar.AppendLine("SELECT		Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Tipo='Comprado' THEN Mon_Bru else 0 END) AS  Mon_Bru1, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Tipo='Comprado' THEN (Mon_Imp1 + Mon_Imp2 + Mon_Imp3) ELSE 0 END) AS  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Tipo<>'Comprado' THEN Mon_Bru ELSE 0 END) AS  Mon_Bru2, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Tipo<>'Comprado' THEN (Mon_Imp1 + Mon_Imp2 + Mon_Imp3) ELSE 0 END) AS  Mon_Imp2, ")

            loComandoSeleccionar.AppendLine("			(SUM(CASE WHEN Tipo='Comprado' THEN Mon_Bru ELSE 0 end) - SUM(CASE WHEN Tipo<>'Comprado' THEN Mon_Bru ELSE 0 END)) AS  Mon_Bru3, ")
            loComandoSeleccionar.AppendLine("			(SUM(CASE WHEN Tipo='Comprado' THEN (Mon_Imp1 + Mon_Imp2 + Mon_Imp3) ELSE 0 END) - SUM(CASE WHEN Tipo<>'Comprado' THEN (Mon_Imp1 + Mon_Imp2 + Mon_Imp3) ELSE 0 END)) AS  Mon_Imp3 ")

            loComandoSeleccionar.AppendLine("FROM		curTemporal ")
            loComandoSeleccionar.AppendLine("GROUP BY	Fec_Ini, Fec_Ini2 ")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCompras_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCompras_Fechas.ReportSource = loObjetoReporte

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
' MAT: 29/04/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
