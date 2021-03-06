﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAuditorias_Expedientes"
'-------------------------------------------------------------------------------------------'
Partial Class fAuditorias_Expedientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
			
			loComandoSeleccionar.AppendLine(" SELECT    Auditorias.Documento, ")
			loComandoSeleccionar.AppendLine(" 			Auditorias.Cod_Usu,")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Auditorias.Registro, 103)       As	Fecha, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN DATEPART(HH, Auditorias.Registro) < 10 THEN '0' + CONVERT(NCHAR(2),DATEPART(HH, Auditorias.Registro))" _
														& " ELSE CONVERT(NCHAR(2),DATEPART(HH, Auditorias.Registro)) END + ':' + CASE WHEN DATEPART(MI, Auditorias.Registro) < 10 THEN '0' + " _
														& "CONVERT(NCHAR(2),DATEPART(MI, Auditorias.Registro)) Else CONVERT(NCHAR(2),DATEPART(MI, Auditorias.Registro)) END AS	Hora, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Tipo, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Tabla, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Cod_Obj, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Notas, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Opcion, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Accion, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Equipo, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Detalle")
            loComandoSeleccionar.AppendLine(" FROM Auditorias,Expedientes ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("			AND	Auditorias.Documento=	Expedientes.Documento ")
            loComandoSeleccionar.AppendLine("			AND	Auditorias.Tabla	=	'Expedientes' ")
            loComandoSeleccionar.AppendLine("			AND	Auditorias.Tipo		=	'Datos' ")
            loComandoSeleccionar.AppendLine("			AND	(Auditorias.Opcion	=	'Expedientes' OR Auditorias.Opcion	=	'Sin opción' OR Auditorias.Opcion	=	'ListarCotizacionesCRM')  ")
			loComandoSeleccionar.AppendLine(" ORDER BY Auditorias.Documento,Auditorias.Registro DESC")
			
        
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

           '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAuditorias_Expedientes", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvfAuditorias_Cotizaciones.ReportSource = loObjetoReporte

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
' MAT: 31/03/2011 : Codigo inicial
'-------------------------------------------------------------------------------------------'
