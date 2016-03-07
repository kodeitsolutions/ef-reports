'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_Numeros
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

			loComandoSeleccionar.Appendline("SELECT     Cuentas_Cobrar.Documento, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Cod_Tip, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Status, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Fec_Ini, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Fec_Fin, "  ) 
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Cod_Cli, "  ) 
			loComandoSeleccionar.Appendline(" 			Clientes.Nom_Cli, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Cod_Ven, "  ) 
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Cod_Tra, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Cod_Mon, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar.Control, "  )
            loComandoSeleccionar.AppendLine("           (CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0")
            loComandoSeleccionar.AppendLine("               WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Bru *(-1)")
            loComandoSeleccionar.AppendLine("               ELSE Cuentas_Cobrar.Mon_Bru")
            loComandoSeleccionar.AppendLine("           END) As Mon_Bru, ")
			loComandoSeleccionar.AppendLine(" 			(CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0")
            loComandoSeleccionar.AppendLine("               ELSE Cuentas_Cobrar.Mon_Imp1")
            loComandoSeleccionar.AppendLine("           END) As Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0")
            loComandoSeleccionar.AppendLine("               WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Net *(-1)")
            loComandoSeleccionar.AppendLine("               ELSE Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("           END) As Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (CASE ")
            loComandoSeleccionar.AppendLine("               WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0")
            loComandoSeleccionar.AppendLine("               WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1)")
            loComandoSeleccionar.AppendLine("               ELSE Cuentas_Cobrar.Mon_Sal")
            loComandoSeleccionar.AppendLine("           END) As Mon_Sal, ")
			loComandoSeleccionar.Appendline(" 			Vendedores.Nom_Ven  "  ) 
			loComandoSeleccionar.Appendline("FROM       Clientes, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Cobrar, "  )
			loComandoSeleccionar.Appendline(" 			Vendedores, "  )
			loComandoSeleccionar.Appendline(" 			Transportes, "  )
			loComandoSeleccionar.Appendline(" 			Monedas "  )
			loComandoSeleccionar.Appendline("WHERE      Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Cobrar.Cod_Tra = Transportes.Cod_Tra "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Cobrar.Cod_Mon = Monedas.Cod_Mon "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven "  )
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status       IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip    BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cla      BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tra      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("		    AND ((" & lcParametro11Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("			OR (" & lcParametro11Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 OR Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Rev      BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro13Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Cobrar.Cod_Tip, Cuentas_Cobrar.Cod_Cli, Cuentas_Cobrar.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY   Cuentas_Cobrar.Cod_Tip," & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_Numeros", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_Numeros.ReportSource =	 loObjetoReporte

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
' CMS: 21/04/09: Estandarización del código y Corrección del estatus
'-------------------------------------------------------------------------------------------'
' CMS: 15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP: 30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS: 03/07/09: Metodo de  Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 13/07/09: Se Agregaron los siguientes filtros: Zona, Tipo de Cliente,
'                 Clase de Cliente.
'                 Verificación de registros
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Multiplicación (*-1) al campo Mon_Net, Mon_Sal, Mon_Bru
'-------------------------------------------------------------------------------------------'
' RJG: 22/07/14: Se eliminó el monto de los documentos anulados para que no entre en los    '
'                totales del reporte.                                                       '
'-------------------------------------------------------------------------------------------'
