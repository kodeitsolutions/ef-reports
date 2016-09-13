Imports System.Data
Partial Class rActivos_Fijos_Ampliado
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


            'Dim lcComandoSelect As String
            Dim loComandoSeleccionar As New StringBuilder()



            '            lcComandoSelect = "SELECT	Cod_Act, " _
            '                   & "Nom_Act, " _
            '& "Status, " _
            '                   & "Fec_Adq, " _
            '                  & "Mon_Adq, " _
            '                 & "Vid_Ano, " _
            '       & "Vid_Mes " _
            '& "(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Activos_Fijos " _
            '    & "FROM	 Activos_Fijos " _
            '    & "WHERE Cod_Act between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
            '       & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
            '    & " ORDER BY Cod_Act, " _
            '       & " Nom_Act "

            loComandoSeleccionar.AppendLine("SELECT		Activos_Fijos.Cod_Act, ")
            loComandoSeleccionar.AppendLine("			Activos_Fijos.Nom_Act, ")
            loComandoSeleccionar.AppendLine("			Activos_Fijos.Fec_Adq, ")
            loComandoSeleccionar.AppendLine("			Activos_Fijos.Vid_Año, ")
            loComandoSeleccionar.AppendLine("			Activos_Fijos.Vid_Mes, ")
            loComandoSeleccionar.AppendLine("			Activos_Fijos.Mon_Adq, ")

            loComandoSeleccionar.AppendLine("			(CASE Activos_Fijos.Status ")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN	'Activo' ")
            loComandoSeleccionar.AppendLine("				WHEN 'S' THEN	'Suspendido' ")
            loComandoSeleccionar.AppendLine("				ELSE			'Inactivo' ")
            loComandoSeleccionar.AppendLine("			END) AS status ")

            loComandoSeleccionar.AppendLine("FROM		Activos_Fijos ")
            loComandoSeleccionar.AppendLine("WHERE		Activos_Fijos.Cod_gru BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 	    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 	    AND Activos_Fijos.Cod_ubi BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 	    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 	    AND Activos_Fijos.Cod_tip BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 	    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 	    AND Activos_Fijos.Cod_cen BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 	    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            Dim loServicios As New cusDatos.goDatos




            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rActivos_Fijos_Ampliado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrActivos_Fijos_Ampliado.ReportSource = loObjetoReporte


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
' MJP   :  15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------' 
' MVP:  05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' PMV:  23/06/15: Creacion del reporte "Activos Fijos Ampliados".
'-------------------------------------------------------------------------------------------'
' EAG:  07/08/15: Agregacion de nuevos parametros
'-------------------------------------------------------------------------------------------'
