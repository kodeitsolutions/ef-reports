'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProposiciones_Prospectos"
'-------------------------------------------------------------------------------------------'

Partial Class rProposiciones_Prospectos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.Appendline("SELECT		Prospectos.Cod_Pro, " )
			loComandoSeleccionar.Appendline("			Prospectos.Nom_Pro, "  )
			loComandoSeleccionar.Appendline("			Prospectos.Status, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Documento, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Fec_ini, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Fec_fin, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Cod_Ven, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Cod_Tra, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Cod_Mon, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Control, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Comentario, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Mon_Net, "  )
			loComandoSeleccionar.Appendline("			Proposiciones.Mon_Sal  "  )
			loComandoSeleccionar.Appendline("FROM		Prospectos, "  )
			loComandoSeleccionar.Appendline("			Proposiciones, "  )
			loComandoSeleccionar.Appendline("			Vendedores, "  )
			loComandoSeleccionar.Appendline("			Transportes, "  )
			loComandoSeleccionar.Appendline("			Monedas "  )
			loComandoSeleccionar.Appendline("WHERE		Prospectos.Cod_Pro = Proposiciones.Cod_Pro "  )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Cod_Tra = Transportes.Cod_Tra "  )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Cod_Mon = Monedas.Cod_Mon "  )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Documento BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro1Hasta  )
			loComandoSeleccionar.Appendline("			AND Prospectos.Cod_Pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro2Hasta )
			loComandoSeleccionar.Appendline("			AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("			AND Proposiciones.Status IN ( " & lcParametro4Desde & ")" )
			loComandoSeleccionar.Appendline("			AND Transportes.Cod_Tra BETWEEN " & lcParametro5Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline("			AND Monedas.Cod_Mon BETWEEN " & lcParametro6Desde )
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Proposiciones.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	Prospectos.Cod_Pro, " & lcOrdenamiento & ", CONVERT(nchar(30), Proposiciones.Fec_Ini,112) DESC, CONVERT(nchar(30), Proposiciones.Fec_Fin,112) DESC")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProposiciones_Prospectos", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProposiciones_Prospectos.ReportSource = loObjetoReporte


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
' MAT: 16/02/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
