'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatos_Reportes"
'-------------------------------------------------------------------------------------------'
Partial Class fDatos_Reportes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
           Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("   SELECT ")
            loComandoSeleccionar.AppendLine("	            Cod_Rep,")
            loComandoSeleccionar.AppendLine("	            Nom_Rep,")
            loComandoSeleccionar.AppendLine("	            Tipo,")
            loComandoSeleccionar.AppendLine("	            Modulo,")
            loComandoSeleccionar.AppendLine("	            Opcion,")
            loComandoSeleccionar.AppendLine("	            Parametro1,")
            loComandoSeleccionar.AppendLine("	            Parametro2,")
            loComandoSeleccionar.AppendLine("	            Parametro3,")
            loComandoSeleccionar.AppendLine("	            Parametro4,")
            loComandoSeleccionar.AppendLine("	            Parametro5,")
            loComandoSeleccionar.AppendLine("	            Parametro6,")
            loComandoSeleccionar.AppendLine("	            Parametro7,")
            loComandoSeleccionar.AppendLine("	            Parametro8,")
            loComandoSeleccionar.AppendLine("	            Parametro9,")
            loComandoSeleccionar.AppendLine("	            Parametro10,")
            loComandoSeleccionar.AppendLine("	            Parametro11,")
            loComandoSeleccionar.AppendLine("	            Parametro12,")
            loComandoSeleccionar.AppendLine("	            Parametro13,")
            loComandoSeleccionar.AppendLine("	            Parametro14,")
            loComandoSeleccionar.AppendLine("	            Parametro15,")
            loComandoSeleccionar.AppendLine("	            Parametro16,")
            loComandoSeleccionar.AppendLine("	            Parametro17,")
            loComandoSeleccionar.AppendLine("	            Parametro18,")
            loComandoSeleccionar.AppendLine("	            Parametro19,")
            loComandoSeleccionar.AppendLine("	            Parametro20,")
            loComandoSeleccionar.AppendLine("	            Orden1,")
            loComandoSeleccionar.AppendLine("	            Orden2,")
            loComandoSeleccionar.AppendLine("               Orden3,")
            loComandoSeleccionar.AppendLine("	            Cam_Ord1,")
            loComandoSeleccionar.AppendLine("	            Cam_Ord2,")
            loComandoSeleccionar.AppendLine("               Cam_Ord3")
            loComandoSeleccionar.AppendLine("   FROM Reportes")
            loComandoSeleccionar.AppendLine("WHERE	")
            loComandoSeleccionar.AppendLine(" 	Cod_Rep between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 	AND 	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 	AND 	Tipo IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 	AND 	Modulo between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 	AND 	" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 	AND 	Opcion between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 	AND 	" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 	AND 	Status IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine(" 	AND 	sistema ='FACTORY " & goAplicacion.pcNombre & "' ")
			loComandoSeleccionar.AppendLine(" 	AND 	GRUPO ='WEB' ")
			loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)
			'loComandoSeleccionar.AppendLine("    " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            cusDatos.goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("fDatos_Reportes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDatos_Reportes.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MJP :  07/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Agregacion de filtro Status
'-------------------------------------------------------------------------------------------'