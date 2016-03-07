'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLlamadas_Usuario"
'-------------------------------------------------------------------------------------------'
Partial Class rLlamadas_Usuario
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_NoFormatear)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_NoFormatear)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT 		")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Unico,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Documento,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Num_Ori,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Num_Des,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Duracion,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Cod_Usu,")
            loComandoSeleccionar.AppendLine(" 			Factory_Global.dbo.Usuarios.Nom_Usu,")
            loComandoSeleccionar.AppendLine(" 			RTRIM(Llamadas.extension)AS extension,")
            loComandoSeleccionar.AppendLine(" 			Llamadas.Comentario")
            loComandoSeleccionar.AppendLine(" FROM	Llamadas		")
            loComandoSeleccionar.AppendLine(" JOIN	Factory_Global.dbo.Usuarios ON 	Factory_Global.dbo.Usuarios.Cod_Usu collate database_default = Llamadas.Cod_Usu	collate database_default")
            loComandoSeleccionar.AppendLine(" WHERE Llamadas.Cod_Usu BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Num_Ori BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Num_Des BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Extension BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Fec_Ini BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Llamadas.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY nom_usu, cod_suc, " & lcOrdenamiento)

            'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLlamadas_Usuario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLlamadas_Usuario.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' EAG: 31/07/15 : Codigo inicial															'
'-------------------------------------------------------------------------------------------'

