'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFacturas_Numeros_SENIAT2"
'-------------------------------------------------------------------------------------------'
Partial Class rFacturas_Numeros_SENIAT2 
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	    Facturas.Documento, ")
            loConsulta.AppendLine(" 			Facturas.Fiscal2, ")
            loConsulta.AppendLine(" 			Facturas.Fec_Ini, ")
            loConsulta.AppendLine(" 			Facturas.Fec_Fin, ")
            loConsulta.AppendLine(" 			Facturas.Cod_Cli, ")
            loConsulta.AppendLine(" 			Clientes.Nom_Cli, ")
            loConsulta.AppendLine(" 			Facturas.Cod_Ven, ")
            loConsulta.AppendLine(" 			Facturas.Cod_Tra, ")
            loConsulta.AppendLine(" 			Facturas.Mon_Imp1, ")
            loConsulta.AppendLine(" 			Vendedores.Status, ")
            loConsulta.AppendLine(" 			Facturas.Control, ")
            loConsulta.AppendLine(" 			Facturas.Mon_Net, ")
            loConsulta.AppendLine(" 			Facturas.Mon_Sal  ")
            loConsulta.AppendLine("FROM         Facturas")
            loConsulta.AppendLine("    JOIN     Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine(" 	JOIN    Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loConsulta.AppendLine(" 			AND Facturas.Documento	BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Cod_Cli	BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Status		IN (" & lcParametro4Desde & ")")
            loConsulta.AppendLine(" 			AND Facturas.Cod_Tra between " & lcParametro5Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Cod_Mon	BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro6Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Cod_For	BETWEEN" & lcParametro7Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro7Hasta)
            loConsulta.AppendLine("             AND Facturas.Cod_Rev    BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("    		    AND " & lcParametro8Hasta)
            loConsulta.AppendLine(" 			AND Facturas.Cod_Suc	BETWEEN" & lcParametro9Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro9Hasta)
            loConsulta.AppendLine("ORDER BY     " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")



            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString, "curReportes")

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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rFacturas_Numeros_SENIAT2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_Numeros_SENIAT2.ReportSource =	 loObjetoReporte	


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
' RJG: 07/10/13: Codigo inicial, a partir de rFacturas_Numeros.                             '
'-------------------------------------------------------------------------------------------'
