Imports System.Data
Imports cusAplicacion

Partial Class rAuditorias_Opciones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    (CASE WHEN Opcion = '' THEN 'Vacío' ELSE Opcion END)    As  Opcion, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Apertura'	THEN 1 ELSE 0 END)) As  Apertura, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Agregar'	THEN 1 ELSE 0 END)) As  Agregar, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Anular'	THEN 1 ELSE 0 END)) As  Anular, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Eliminar'	THEN 1 ELSE 0 END)) As  Eliminar, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Modificar'	THEN 1 ELSE 0 END)) As  Modificar, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Reporte'	THEN 1 ELSE 0 END)) As  Reporte, ")
            loComandoSeleccionar.AppendLine("           SUM(1) As TotalOpcion ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpOpciones ")
            loComandoSeleccionar.AppendLine(" FROM      Auditorias ")
            loComandoSeleccionar.AppendLine("           WHERE Registro  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cod_Usu     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Tabla       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Opcion      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Documento   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Codigo      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Cod_Emp     Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  Auditorias.Opcion ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Auditorias.Opcion ")

			If lcParametro7Desde = 0 Then
				loComandoSeleccionar.AppendLine(" SELECT  Opcion, ")
			Else
				loComandoSeleccionar.AppendLine(" SELECT  TOP " & lcParametro7Desde & "  Opcion, ")
			End If
            
            loComandoSeleccionar.AppendLine("           Apertura, ")
            loComandoSeleccionar.AppendLine("           Agregar, ")
            loComandoSeleccionar.AppendLine("           Anular, ")
            loComandoSeleccionar.AppendLine("           Eliminar, ")
            loComandoSeleccionar.AppendLine("           Modificar, ")
            loComandoSeleccionar.AppendLine("           Reporte, ")
            loComandoSeleccionar.AppendLine("           TotalOpcion ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpOpciones ")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'loComandoSeleccionar.AppendLine(" ORDER BY  TotalOpcion DESC ")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_Opciones", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_Opciones.ReportSource = loObjetoReporte

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
' JJD: 23/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 10/01/09: Se ajusto el codigo a la sintaxis. Se ordeno por la mejor opcion.
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
' JJD: 26/04/10: Se incluyo el filtro Los Primeros
'-------------------------------------------------------------------------------------------'
