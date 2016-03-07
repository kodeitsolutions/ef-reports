'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAuditorias_OpcionesUsuario"
'-------------------------------------------------------------------------------------------'
Partial Class rAuditorias_OpcionesUsuario
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
            loComandoSeleccionar.AppendLine("           SUM(1) As TotalOpcion, ")
            loComandoSeleccionar.AppendLine("           Cod_Usu ")
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
            loComandoSeleccionar.AppendLine(" GROUP BY  Opcion, Cod_Usu ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Usu, TotalOpcion Desc")

			
			loComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER (PARTITION BY Cod_Usu  ORDER BY Cod_Usu ) AS ROW_NUMBER,  ")
			loComandoSeleccionar.AppendLine("           #tmpOpciones.Opcion, ")			
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Apertura, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Agregar, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Anular, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Eliminar, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Modificar, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Reporte, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.TotalOpcion, ")
            loComandoSeleccionar.AppendLine("           #tmpOpciones.Cod_Usu ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpOpciones2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpOpciones ")
            
			loComandoSeleccionar.AppendLine(" SELECT	#tmpOpciones2.ROW_NUMBER,")
			loComandoSeleccionar.AppendLine("       	#tmpOpciones2.Opcion,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Apertura,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Agregar,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Anular,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Eliminar,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Modificar,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.Reporte,  ")
			loComandoSeleccionar.AppendLine("            #tmpOpciones2.TotalOpcion, ")
			loComandoSeleccionar.AppendLine(" 		   Usuarios.Cod_Usu, ")
			loComandoSeleccionar.AppendLine(" 		   Usuarios.Nom_Usu  ")
			loComandoSeleccionar.AppendLine("  FROM      #tmpOpciones2  ")
			loComandoSeleccionar.AppendLine("  JOIN Factory_Global.dbo.Usuarios AS USuarios ON Usuarios.Cod_Usu collate Modern_Spanish_CI_AS = #tmpOpciones2.Cod_Usu collate Modern_Spanish_CI_AS")
			
			If lcParametro7Desde > 0
				loComandoSeleccionar.AppendLine("  WHERE #tmpOpciones2.ROW_NUMBER < = " & lcParametro7Desde)
			End If					           
            
            loComandoSeleccionar.AppendLine(" ORDER BY Usuarios.Cod_Usu, " & lcOrdenamiento & ", #tmpOpciones2.ROW_NUMBER")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_OpcionesUsuario", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_OpcionesUsuario.ReportSource = loObjetoReporte

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
' CMS: 27/04/10: Codigo inicial
'-------------------------------------------------------------------------------------------'