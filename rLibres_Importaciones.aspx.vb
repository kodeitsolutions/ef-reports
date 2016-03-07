'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibres_Importaciones"
'-------------------------------------------------------------------------------------------'
Partial Class rLibres_Importaciones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT		Libres_Importaciones.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Tasa, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Status, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Mon_Imp1 AS Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Importaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Monedas.Nom_Mon ")
            loComandoSeleccionar.AppendLine(" From			Libres_Importaciones, ")
            loComandoSeleccionar.AppendLine("				Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE			Libres_Importaciones.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 				AND Libres_Importaciones.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Libres_Importaciones.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Libres_Importaciones.Status IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Libres_Importaciones.cod_rev between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY      Libres_Importaciones." & lcOrdenamiento)

            'loComandoSeleccionar.AppendLine(" ORDER BY		Libres_Importaciones.Documento ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibres_Importaciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibres_Importaciones.ReportSource = loObjetoReporte

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
' CMS: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'
