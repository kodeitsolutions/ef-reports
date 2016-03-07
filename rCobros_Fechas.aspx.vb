'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Fechas
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

			loComandoSeleccionar.AppendLine("SELECT			Cobros.Documento			AS Documento,	") 
			loComandoSeleccionar.AppendLine(" 				Cobros.Fec_Ini				AS Fec_Ini,		")
			loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Cli				AS Cod_Cli,		")
			loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli			AS Nom_Cli,		") 
			loComandoSeleccionar.AppendLine(" 				Cobros.Mon_Net				AS Mon_Net,		")
			loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Mon				AS Cod_Mon,		")
			loComandoSeleccionar.AppendLine(" 				Cobros.Cod_Ven				AS Cod_Ven,		") 
			loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven			AS Nom_Ven,		")
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Tip_Ope		AS Tip_Ope,		")
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Renglon		AS Renglon,		") 
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Mon_Net		AS Cob_Ope,		")
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Doc_Des		AS Doc_Des,		")
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Num_Doc		AS Num_Doc,		") 
			loComandoSeleccionar.AppendLine(" 				CASE WHEN(Detalles_Cobros.Cod_Cue = '')		")
			loComandoSeleccionar.AppendLine(" 					THEN	Detalles_Cobros.Cod_Ban			") 
			loComandoSeleccionar.AppendLine(" 					ELSE	Cuentas_Bancarias.Cod_Ban		")
			loComandoSeleccionar.AppendLine(" 				END							AS Cod_Ban,		")
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Cod_Caj		AS Cod_Caj,		") 
			loComandoSeleccionar.AppendLine(" 				Detalles_Cobros.Cod_Cue		AS Cod_Cue		")
			loComandoSeleccionar.AppendLine("FROM			Cobros " )
			loComandoSeleccionar.AppendLine("	JOIN 		Detalles_Cobros ON Cobros.Documento = Detalles_Cobros.Documento" )
			loComandoSeleccionar.AppendLine("	JOIN 		Clientes ON Cobros.Cod_Cli = Clientes.Cod_Cli" )
			loComandoSeleccionar.AppendLine("	JOIN 		Vendedores ON Cobros.Cod_Ven = Vendedores.Cod_Ven " )
			loComandoSeleccionar.AppendLine("	LEFT JOIN	Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Detalles_Cobros.Cod_Cue" )
			loComandoSeleccionar.AppendLine("WHERE		Cobros.Documento BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cobros.Fec_Ini BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Cli BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Mon BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Ven BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("       	AND Cobros.Cod_Rev BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       	 AND Cobros.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Cobros.Fec_Ini , " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
			
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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCobros_Fechas", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Fechas.ReportSource =	 loObjetoReporte

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
' GCR: 09/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  16/07/09: Metodo de ordenamiento, verificacion de registros
'-------------------------------------------------------------------------------------------'
' MAT:  28/06/11: Ajuste del Select, Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG:  21/03/12: Se agregó el código de cuenta, y el banco asociado a la misma (si aplica).'
'-------------------------------------------------------------------------------------------'
' RJG:  12/04/12: Se corrigiéton los total de registros: estaban invertidos.				'
'-------------------------------------------------------------------------------------------'
