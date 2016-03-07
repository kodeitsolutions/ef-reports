'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rExpedientes"
'-------------------------------------------------------------------------------------------'
Partial Class rExpedientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()
			loComandoSeleccionar.AppendLine(" SELECT")
			loComandoSeleccionar.AppendLine(" 	ex.documento	AS documento,")
			loComandoSeleccionar.AppendLine(" 	ex.expediente	AS expediente,")
			loComandoSeleccionar.AppendLine(" 	ex.tipo			AS tipo,")
			loComandoSeleccionar.AppendLine(" 	ex.clase		AS clase,")
			loComandoSeleccionar.AppendLine(" 	ex.status		AS status,")
			loComandoSeleccionar.AppendLine(" 	ex.importacion	AS importacion,")
			loComandoSeleccionar.AppendLine(" 	ex.cod_log		AS cod_log,")
			loComandoSeleccionar.AppendLine(" 	ex.nom_log		AS nom_log,")
			loComandoSeleccionar.AppendLine(" 	ex.fec_ini		AS fec_ini,")
			loComandoSeleccionar.AppendLine(" 	ex.fec_fin		AS fec_fin,")
			loComandoSeleccionar.AppendLine(" 	ex.fec_rec		AS fec_rec,")
			loComandoSeleccionar.AppendLine(" 	ex.seguimientos AS seguimientos,")
			loComandoSeleccionar.AppendLine(" 	ex.comentario	AS comentario,")
			loComandoSeleccionar.AppendLine(" 	ex.prioridad	AS prioridad,")
			loComandoSeleccionar.AppendLine(" 	ex.nivel		AS nivel")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" FROM expedientes AS ex")
			loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" WHERE     ")
            loComandoSeleccionar.AppendLine(" 			documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Status IN ( " & lcParametro1Desde & " ) ")   
            loComandoSeleccionar.AppendLine(" 			AND Cod_Log BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Fec_Ini BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Nom_Suc ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rExpedientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrExpedientes.ReportSource = loObjetoReporte

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
' BAP: 05/08/10: Programacion inicial
'-------------------------------------------------------------------------------------------'