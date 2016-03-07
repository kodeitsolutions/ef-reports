'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDescuentos_cVolumen"
'-------------------------------------------------------------------------------------------'
Partial Class rDescuentos_cVolumen
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            
			Dim lcParametro0Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			'Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

			Dim lcOrdenacion As String = goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		DA.Clase									AS Clase,")
			loComandoSeleccionar.AppendLine("			CASE DA.Clase")
			loComandoSeleccionar.AppendLine("				WHEN 'Articulo'		THEN DA.Cod_Art")
			loComandoSeleccionar.AppendLine("				WHEN 'Departamento' THEN DA.Cod_Dep")
			loComandoSeleccionar.AppendLine("				WHEN 'Segmento'		THEN DA.Cod_Seg")
			loComandoSeleccionar.AppendLine("			END											AS Cod_Reg,")
			loComandoSeleccionar.AppendLine("			CASE DA.Clase")
			loComandoSeleccionar.AppendLine("				WHEN 'Articulo'		THEN A.Nom_Art")
			loComandoSeleccionar.AppendLine("				WHEN 'Departamento' THEN D.Nom_Dep")
			loComandoSeleccionar.AppendLine("				WHEN 'Segmento'		THEN S.Nom_Seg")
			loComandoSeleccionar.AppendLine("			END											AS Nom_Reg,")
			loComandoSeleccionar.AppendLine("			DA.Status									AS Status,")
			loComandoSeleccionar.AppendLine("			DA.Tip_Cli									AS Tip_Cli,")
			loComandoSeleccionar.AppendLine("			DA.Can_Des									AS Can_Des,")
			loComandoSeleccionar.AppendLine("			DA.Can_Has									AS Can_Has,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des1									AS Por_Des1")
			loComandoSeleccionar.AppendLine("FROM		Descuentos_Articulos AS DA")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Articulos AS A")
			loComandoSeleccionar.AppendLine("		ON	A.Cod_Art = DA.Cod_art")
			loComandoSeleccionar.AppendLine("		AND	A.Cod_Art BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("	LEFT JOIN Departamentos AS D")
			loComandoSeleccionar.AppendLine("		ON	D.Cod_Dep = DA.Cod_Dep")
			loComandoSeleccionar.AppendLine("		AND	D.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("	LEFT JOIN Segmentos AS S")
			loComandoSeleccionar.AppendLine("		ON	S.Cod_Seg = DA.Cod_Seg")
			loComandoSeleccionar.AppendLine("		AND	S.Cod_Seg BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("WHERE		DA.Adicional = 'Volumen'")
            loComandoSeleccionar.AppendLine("		AND	DA.Status IN (" & lcParametro0Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	DA.Clase IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	DA.Tip_Cli BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenacion & ", Cod_Reg, DA.Tip_Cli, DA.Can_Des")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDescuentos_cVolumen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDescuentos_cVolumen.ReportSource = loObjetoReporte

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
