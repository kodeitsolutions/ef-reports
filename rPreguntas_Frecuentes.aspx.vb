'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPreguntas_Frecuentes"
'-------------------------------------------------------------------------------------------'
Partial Class rPreguntas_Frecuentes

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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		Cod_Pre, ")
            loComandoSeleccionar.AppendLine(" 		Nom_Pre, ")
            loComandoSeleccionar.AppendLine(" CASE ")
            loComandoSeleccionar.AppendLine(" 		WHEN Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 		WHEN Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 		WHEN Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" END AS Status, ")
            loComandoSeleccionar.AppendLine(" 		Pregunta, ")
            loComandoSeleccionar.AppendLine(" 		Respuesta, ")
            loComandoSeleccionar.AppendLine(" 		Grupo, ")
            loComandoSeleccionar.AppendLine(" 		Clase, ")
            loComandoSeleccionar.AppendLine(" 		Tipo, ")
            loComandoSeleccionar.AppendLine(" 		Cod_Art ")
            loComandoSeleccionar.AppendLine(" FROM Preguntas_Frecuentes")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Cod_Pre                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cod_Art                  Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("       And Tipo = " & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("       And Clase = " & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("       And Grupo = " & lcParametro5Desde)
            End If

            if lcOrdenamiento.ToUpper.Replace("ASC","").Trim = "PREGUNTA" or lcOrdenamiento.ToUpper.Replace("DESC","").Trim = "PREGUNTA"then 
				loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
			else
				loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento & ", Tipo, Clase, Grupo")
			End If

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
							  
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPreguntas_Frecuentes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrPreguntas_Frecuentes.ReportSource = loObjetoReporte

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
' CMS: 19/02/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'