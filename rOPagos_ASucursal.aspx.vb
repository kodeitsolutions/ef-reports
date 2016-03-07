'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_ASucursal"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_ASucursal

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Ordenes_Pagos.Cod_Suc,  ")
            loComandoSeleccionar.AppendLine(" 				Sucursales.Nom_Suc,   ")
            loComandoSeleccionar.AppendLine(" 				SUM(Renglones_oPagos.Mon_deb)	AS Mon_Deb,   ")
            loComandoSeleccionar.AppendLine(" 				SUM(Renglones_oPagos.Mon_hab)	AS Mon_Hab,   ")
            loComandoSeleccionar.AppendLine(" 				SUM(Renglones_oPagos.Mon_imp1)	AS Mon_Imp1   ")
            loComandoSeleccionar.AppendLine(" FROM			Ordenes_Pagos,    ")
            loComandoSeleccionar.AppendLine(" 				Renglones_oPagos,    ")
            loComandoSeleccionar.AppendLine(" 				Sucursales	   ")
            loComandoSeleccionar.AppendLine(" WHERE			Ordenes_Pagos.Documento = Renglones_oPagos.Documento  ")
            loComandoSeleccionar.AppendLine(" 				AND Ordenes_Pagos.Cod_Suc = Sucursales.Cod_Suc  ")

            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Pro   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Suc  Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Mon  Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Mon  Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_rev  Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY		Ordenes_Pagos.Cod_Suc, Sucursales.Nom_Suc   ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_ASucursal", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_ASucursal.ReportSource = loObjetoReporte

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
' CMS: 18/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 25/10/10: Mejora de la parte visual del reporte
'-------------------------------------------------------------------------------------------'

