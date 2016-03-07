Imports System.Data
Imports cusAplicacion

Partial Class rAuditorias_UsuariosFechas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT    (CASE WHEN Cod_Usu = '' THEN 'Vacio' ELSE Cod_Usu END)	AS	Cod_Usu, ")
            lcComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Registro, 112)					    AS	Fecha2,	")
            lcComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Registro, 103)                       AS	Fecha, ")
            lcComandoSeleccionar.AppendLine("           Accion,")
            lcComandoSeleccionar.AppendLine("           dbo.udf_GetISOWeekDay(Registro)         AS  Dias  ")
            lcComandoSeleccionar.AppendLine(" INTO      #tmpAuditorias ")
            lcComandoSeleccionar.AppendLine(" FROM      Auditorias ")
            lcComandoSeleccionar.AppendLine(" WHERE     Registro        Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cod_Usu     Between " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Tabla       Between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Opcion      Between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("           And Documento   Between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Codigo      Between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("           And Cod_Emp     Between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)

            lcComandoSeleccionar.AppendLine(" SELECT    Cod_Usu, ")
            lcComandoSeleccionar.AppendLine("           Fecha, ")
            lcComandoSeleccionar.AppendLine("           Fecha2, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Apertura'	THEN 1 ELSE 0 END)) As Apertura, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Agregar'	THEN 1 ELSE 0 END)) As Agregar, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Anular'	THEN 1 ELSE 0 END)) As Anular, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Eliminar'	THEN 1 ELSE 0 END)) As Eliminar, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Modificar'	OR Accion = 'Confirmar'	OR Accion = 'Revertir'	THEN 1 ELSE 0 END)) As Modificar, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = 'Reporte'	OR Accion = 'Formato'	THEN 1 ELSE 0 END)) As Reporte, ")
            lcComandoSeleccionar.AppendLine("           SUM((CASE WHEN Accion = ''	        THEN 1 ELSE 0 END)) As Vacio, ")
            lcComandoSeleccionar.AppendLine("           SUM(1) As TotalUsuario, ")
            lcComandoSeleccionar.AppendLine("           SUM(1) As TotalUsuario, ")
            lcComandoSeleccionar.AppendLine("           (CASE Dias ")
            lcComandoSeleccionar.AppendLine("                 WHEN 1 THEN 'LUNES'")
            lcComandoSeleccionar.AppendLine("                 WHEN 2 THEN 'MARTES'")
            lcComandoSeleccionar.AppendLine("                 WHEN 3 THEN 'MIERCOLES'")
            lcComandoSeleccionar.AppendLine("                 WHEN 4 THEN 'JUEVES'")
            lcComandoSeleccionar.AppendLine("                 WHEN 5 THEN 'VIERNES'")
            lcComandoSeleccionar.AppendLine("                 WHEN 6 THEN 'SABADO'")
            lcComandoSeleccionar.AppendLine("                 WHEN 7 THEN 'DOMINGO'")
            lcComandoSeleccionar.AppendLine("           END) AS Dia")
            lcComandoSeleccionar.AppendLine(" FROM      #tmpAuditorias")
            lcComandoSeleccionar.AppendLine(" GROUP BY  Cod_Usu, Fecha2, Fecha, Dias ")
            lcComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)


            'lcComandoSeleccionar.AppendLine(" ORDER BY  Cod_Usu, Fecha2 ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_UsuariosFechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_UsuariosFechas.ReportSource = loObjetoReporte

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
' JJD: 13/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 26/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' EAG: 28/08/15: Se agrego el campo dia.                          							'
'-------------------------------------------------------------------------------------------'

