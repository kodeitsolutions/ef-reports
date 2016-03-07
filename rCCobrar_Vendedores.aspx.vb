'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_Vendedores
    Inherits vis2Formularios.frmReporte
    
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
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine( "  SELECT		Cuentas_Cobrar.Documento, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Cod_Tip, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Fec_Ini, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Fec_Fin, " ) 
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Cod_Cli, " ) 
			loComandoSeleccionar.AppendLine( " 				Clientes.Nom_Cli, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Cod_Ven, " ) 
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Cod_Tra, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Cod_Mon, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Control, " )
            'loComandoSeleccionar.AppendLine( " 			Cuentas_Cobrar.Mon_Bru, " )
            loComandoSeleccionar.AppendLine("               (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Bru *(-1) Else Cuentas_Cobrar.Mon_Bru End) As Mon_Bru, ")
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar.Mon_Imp1, " )
            'loComandoSeleccionar.AppendLine( " 			Cuentas_Cobrar.Mon_Net, " )
            loComandoSeleccionar.AppendLine("               (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Net *(-1) Else Cuentas_Cobrar.Mon_Net End) As Mon_Net, ")
            loComandoSeleccionar.AppendLine("               (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End) As Mon_Sal,  ")
			loComandoSeleccionar.AppendLine( " 				Vendedores.Nom_Ven  " ) 
			loComandoSeleccionar.AppendLine( " FROM			Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar, " )
			loComandoSeleccionar.AppendLine( " 				Vendedores, " )
			loComandoSeleccionar.AppendLine( " 				Transportes, " )
			loComandoSeleccionar.AppendLine( " 				Monedas " )
			loComandoSeleccionar.AppendLine( " WHERE		Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli " )
            loComandoSeleccionar.AppendLine("               AND 	Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("               AND 	Cuentas_Cobrar.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("               AND 	Cuentas_Cobrar.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine("               AND 	Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Ven      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Zon      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Status       IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Tip    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cla      Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tra      Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Mon      Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("		    AND ((" & lcParametro11Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("			OR (" & lcParametro11Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      Between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Rev      Between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro13Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Cuentas_Cobrar.Cod_Ven,  Cuentas_Cobrar.Cod_Cli,  Cuentas_Cobrar.Cod_Tip, Cuentas_Cobrar.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY       Cuentas_Cobrar.Cod_Ven," & lcOrdenamiento)

         
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCCobrar_Vendedores", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_Vendedores.ReportSource =	 loObjetoReporte

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
' JJD: 22/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  13/07/09: Se Agregaron los siguientes filtros: Zona, Tipo de Cliente,
'                 Clase de Cliente, Revisión.
'                 Verificación de registros
'-------------------------------------------------------------------------------------------'
' CMS:  15/07/09: Multiplicación (*-1) al campo Mon_Net, Mon_Sal, Mon_Bru
'-------------------------------------------------------------------------------------------'