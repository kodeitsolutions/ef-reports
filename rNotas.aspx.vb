﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNotas"
'-------------------------------------------------------------------------------------------'
Partial Class rNotas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		Nom_Not, ")
            loComandoSeleccionar.AppendLine(" 		Case ")
            loComandoSeleccionar.AppendLine(" 			When Color = '#8C4AE6' Then 'Azul Claro'")
            loComandoSeleccionar.AppendLine(" 			When Color = '#003399' Then 'Azul Oscuro'")
            loComandoSeleccionar.AppendLine(" 			When Color = '#13920D' Then 'Verde'")
            loComandoSeleccionar.AppendLine(" 			When Color = '#FFCC00' Then 'Amarillo'")
            loComandoSeleccionar.AppendLine(" 			When Color = '#FF9933' Then 'Anaranjado'")
            loComandoSeleccionar.AppendLine(" 			When Color = '#E35C2F' Then 'Rojo'")
            loComandoSeleccionar.AppendLine(" 		END AS Color, ")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 				WHEN Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" 		END AS Status, ")
            loComandoSeleccionar.AppendLine(" 		Comentario, ")
            loComandoSeleccionar.AppendLine(" 		Fec_Ini ")
            loComandoSeleccionar.AppendLine(" FROM NOtas")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            
            If Trim(cusAplicacion.goReportes.paParametrosIniciales(0)) <> "" Then

	            loComandoSeleccionar.AppendLine("           Nom_Not    LIKE '%" & cusAplicacion.goReportes.paParametrosIniciales(0) & "%' And")
	        
	        End If
	        
            loComandoSeleccionar.AppendLine("           Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Cod_Usu = '" & goUsuario.pcCodigo.ToString & "'")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNotas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrNotas.ReportSource = loObjetoReporte

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
' CMS: 20/07/2010: Codigo inicial.
'-------------------------------------------------------------------------------------------'