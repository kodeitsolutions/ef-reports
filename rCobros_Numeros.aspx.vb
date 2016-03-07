'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Numeros
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
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Cobros.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("				Cobros.Mon_Net,				")
            loComandoSeleccionar.AppendLine("				Cobros.Mon_Des,				")
            loComandoSeleccionar.AppendLine("				Cobros.Mon_Ret,				")
            loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Cobros.Renglon, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Cobros.Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Cobros.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("               (	CASE				")
            loComandoSeleccionar.AppendLine("						WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' OR Cobros.Automatico = 1")
            loComandoSeleccionar.AppendLine("						THEN Renglones_Cobros.Mon_Abo ")
            loComandoSeleccionar.AppendLine("						ELSE (Renglones_Cobros.Mon_Abo * -1)")
            loComandoSeleccionar.AppendLine("					END)  AS  Cargo, ")
            loComandoSeleccionar.AppendLine("               (	CASE				")
            loComandoSeleccionar.AppendLine("						WHEN (Cuentas_Cobrar.Tip_Doc = 'Debito' OR Cobros.Automatico = 1)")
            loComandoSeleccionar.AppendLine("						THEN Renglones_Cobros.Mon_Net ")
            loComandoSeleccionar.AppendLine("						ELSE (Renglones_Cobros.Mon_Net * -1)")
            loComandoSeleccionar.AppendLine("					END)  AS  Mon_Doc")
            loComandoSeleccionar.AppendLine("FROM			Cobros")
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Cobros ")
            loComandoSeleccionar.AppendLine("			ON	Cobros.Documento = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("	JOIN		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("			ON	Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip ")
            loComandoSeleccionar.AppendLine("	JOIN		Clientes")
            loComandoSeleccionar.AppendLine("			ON	Clientes.Cod_Cli = Cobros.Cod_Cli")
            loComandoSeleccionar.AppendLine("	JOIN		Vendedores")
            loComandoSeleccionar.AppendLine("			ON	Vendedores.Cod_Ven = Cobros.Cod_Ven")
            loComandoSeleccionar.AppendLine("	JOIN		Monedas ")
            loComandoSeleccionar.AppendLine("			ON	Monedas.Cod_Mon = Cobros.Cod_Mon")
            loComandoSeleccionar.AppendLine(" WHERE			Cobros.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cobros.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cobros.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Monedas.Cod_Mon BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Ven BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cobros.Status  IN(" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("					AND Cobros.Cod_Rev BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("					AND Cobros.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY       Cobros.Documento, " & lcOrdenamiento )

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCobros_Numeros", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrCobros_Numeros.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 22/09/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' CMS: 15/05/09: Filtro “Revisión:”															'
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro “Sucursal:”															'
'-------------------------------------------------------------------------------------------'
' CMS: 31/08/09: Según la naturaleza del documento de multiplica por -1 los siguientes campo:'
'                Cargo, Mon_Doc, Mon_Net													'
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/12: Corrección en el signo de los cobros de Adelandos a Clientes.				'
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/12: Se agegaron los totales cobrados (cobros.mon_net), de descuentos y de		'
'				 retenciones.																'
'-------------------------------------------------------------------------------------------'
' RJG: 21/03/12: Se agegó el monto abonado a la última columna del reporte (Cargo).			'
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de documentos.											'
'-------------------------------------------------------------------------------------------'
