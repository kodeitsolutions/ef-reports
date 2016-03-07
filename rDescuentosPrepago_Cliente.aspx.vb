'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDescuentosPrepago_Cliente"
'-------------------------------------------------------------------------------------------'
Partial Class rDescuentosPrepago_Cliente

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
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
            Dim lcParametro9Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            
            Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		Clientes.Origen				AS Origen,")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cli					AS Cod_Cli,")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli			AS Nom_Cli,")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cla					AS Cod_Cla,")
			loComandoSeleccionar.AppendLine("			CA.Nom_Cla					AS Nom_Cla,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des1					AS Por_Des1,")
			loComandoSeleccionar.AppendLine("			CASE DA.Status")
			loComandoSeleccionar.AppendLine("				WHEN 'A'	THEN 'Activo'")
			loComandoSeleccionar.AppendLine("				WHEN 'I'	THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine("				WHEN 'S'	THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine("				ELSE			 'Desconocido'")
			loComandoSeleccionar.AppendLine("			END							AS Status")
			loComandoSeleccionar.AppendLine("FROM		Clientes ")
			loComandoSeleccionar.AppendLine("	JOIN	Descuentos_Articulos AS DA")
			loComandoSeleccionar.AppendLine("		ON	DA.Cod_Cli		= Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine("		AND	DA.Origen		= 'Clientes'")
			loComandoSeleccionar.AppendLine("		AND	DA.Adicional	= 'Prepago'")
			loComandoSeleccionar.AppendLine("		AND	DA.Clase		= 'Clase'")
			loComandoSeleccionar.AppendLine("		AND	DA.Status IN (" & lcParametro9Desde & ")")
			loComandoSeleccionar.AppendLine("	JOIN	Clases_Articulos AS CA ")
			loComandoSeleccionar.AppendLine("		ON	 CA.Cod_Cla = DA.Cod_Cla")
			loComandoSeleccionar.AppendLine("WHERE		Clientes.Cod_Cli BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("		AND	Clientes.Tip_Cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Cla BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Ven BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Zon BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Suc BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Can BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("		AND	Clientes.Cod_Ope BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento & ", DA.Cod_Cla")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDescuentosPrepago_Cliente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDescuentosPrepago_Cliente.ReportSource = loObjetoReporte

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
' RJG: 14/05/12: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
