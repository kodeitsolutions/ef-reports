'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTickets_Resumido_Fonax"
'-------------------------------------------------------------------------------------------'
Partial Class rTickets_Resumido_Fonax 
    Inherits vis2Formularios.frmReporte
    
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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT	MAX(Auditorias.Registro) AS Fecha_Confirmacion, "  )
			loComandoSeleccionar.Appendline("		Auditorias.Documento "  )
			loComandoSeleccionar.Appendline("INTO #tmpTemporal "  )
			loComandoSeleccionar.Appendline("FROM Auditorias "  )
			loComandoSeleccionar.Appendline("WHERE 	Auditorias.Tabla = 'Cotizaciones' "  )
			loComandoSeleccionar.Appendline("		AND Auditorias.Accion = 'Confirmar' "  )
			loComandoSeleccionar.Appendline("GROUP BY Auditorias.Documento "  )
			loComandoSeleccionar.Appendline("ORDER BY Documento ASC"  )
			
			
			loComandoSeleccionar.Appendline("SELECT	Cotizaciones.Documento, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Fec_Ini AS Fecha_Emision, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Reg_Ini AS Fecha_Creacion, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Comentario, ")
			loComandoSeleccionar.Appendline("		Vendedores.Nom_Ven, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Cli,"  )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli, "  )
			loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Cotizaciones.Status = 'Pendiente' THEN (DATEDIFF(Day,Cotizaciones.Reg_Ini,getDate()))")
			loComandoSeleccionar.AppendLine("			WHEN Cotizaciones.Status = 'Confirmado' THEN (DATEDIFF(Day,Cotizaciones.Reg_Ini,#tmpTemporal.Fecha_Confirmacion))")
			loComandoSeleccionar.AppendLine("			ELSE 0")
			loComandoSeleccionar.AppendLine("		END AS Dias_Activo")
			loComandoSeleccionar.AppendLine("FROM Cotizaciones")
			loComandoSeleccionar.AppendLine("JOIN Clientes ON (Cotizaciones.Cod_Cli = Clientes.Cod_Cli )")
			loComandoSeleccionar.AppendLine("JOIN Vendedores ON (Cotizaciones.Cod_Ven = Vendedores.Cod_Ven)")
			loComandoSeleccionar.AppendLine("LEFT JOIN #tmpTemporal ON (#tmpTemporal.Documento = Cotizaciones.Documento)")
			loComandoSeleccionar.Appendline("WHERE	Cotizaciones.Documento Between " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Fec_Ini between " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine("		And Cotizaciones.Status IN (" & lcParametro2Desde & ")")
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Cli between " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Ven between " & lcParametro4Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro4Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Tra between " & lcParametro5Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Rev between " & lcParametro6Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro6Hasta )
             loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString , "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTickets_Resumido_Fonax", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrTickets_Resumido_Fonax.ReportSource =	 loObjetoReporte	

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
' MAT: 20/06/11: Código Inicial
'-------------------------------------------------------------------------------------------'