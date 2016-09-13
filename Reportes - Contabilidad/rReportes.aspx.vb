'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReportes"
'-------------------------------------------------------------------------------------------'
Partial Class rReportes
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


            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		    Cod_Rep, ")
            loComandoSeleccionar.AppendLine(" 		    Nom_Rep, ")
            loComandoSeleccionar.AppendLine(" 		    Opcion, ")
            loComandoSeleccionar.AppendLine(" 		    Archivo ")
            loComandoSeleccionar.AppendLine(" FROM ")
            loComandoSeleccionar.AppendLine(" 		    Reportes ")
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

           
            
            Dim loServicios As New cusDatos.goDatos

            cusDatos.goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReportes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrReportes.ReportSource = loObjetoReporte

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
