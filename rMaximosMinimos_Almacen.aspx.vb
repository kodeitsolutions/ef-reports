'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMaximosMinimos_Almacen"
'-------------------------------------------------------------------------------------------'
Partial Class rMaximosMinimos_Almacen
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
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
			Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
			Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
			Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
			Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))

			Dim lcOrdenacion As String = goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		Sucursales.Cod_Suc		AS Cod_Suc,")
			loComandoSeleccionar.AppendLine("			Sucursales.Nom_Suc		AS Nom_Suc,")
			loComandoSeleccionar.AppendLine("			Almacenes.Cod_Alm		AS Cod_Alm,")
			loComandoSeleccionar.AppendLine("			Almacenes.Nom_Alm		AS Nom_Alm,")
			loComandoSeleccionar.AppendLine("			Articulos.Cod_Art		AS Cod_Art,")
			loComandoSeleccionar.AppendLine("			Articulos.Nom_Art		AS Nom_Art,")
			loComandoSeleccionar.AppendLine("			SA.Exi_Min				AS Exi_Min,")
			loComandoSeleccionar.AppendLine("			SA.Exi_Max				AS Exi_Max,")
			loComandoSeleccionar.AppendLine("			SA.Exi_Pto				AS Exi_Pto")
			loComandoSeleccionar.AppendLine("FROM		Sucursales_Articulos AS SA")
			loComandoSeleccionar.AppendLine("	JOIN	Sucursales ")
			loComandoSeleccionar.AppendLine("		ON	Sucursales.Cod_Suc = SA.Cod_Suc")
            loComandoSeleccionar.AppendLine("		AND	Sucursales.Cod_Suc BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("	JOIN	Almacenes ")
			loComandoSeleccionar.AppendLine("		ON	Almacenes.Cod_Alm = SA.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND	Almacenes.Cod_Alm BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("	JOIN	Articulos ")
			loComandoSeleccionar.AppendLine("		ON	Articulos.Cod_Art = SA.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Art BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Sec BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Pro BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Mar BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Tip BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Cla BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Ubi BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("		AND	Articulos.Cod_Seg BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenacion & ", Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMaximosMinimos_Almacen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMaximosMinimos_Almacen.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 03/04/12: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
