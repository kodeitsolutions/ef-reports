'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRetenciones_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rRetenciones_Proveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("DECLARE @lnCero AS Decimal(28, 10);")
			loComandoSeleccionar.AppendLine("SET @lnCero = 0;")
			loComandoSeleccionar.AppendLine("SELECT		")
			loComandoSeleccionar.AppendLine("  			Proveedores.Cod_Pro								AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("  			Proveedores.Nom_Pro								AS Nom_Pro,")
			loComandoSeleccionar.AppendLine("  			CASE WHEN Cuentas_Pagar.Cod_Tip = 'RETIVA' ")
			loComandoSeleccionar.AppendLine("  				THEN Cuentas_Pagar.Mon_Net")  
			loComandoSeleccionar.AppendLine("  				Else @lnCero")				  
			loComandoSeleccionar.AppendLine("  			END												AS Impuesto,")
			loComandoSeleccionar.AppendLine("  			CASE WHEN Cuentas_Pagar.Cod_Tip = 'ISLR' ")
			loComandoSeleccionar.AppendLine("  				THEN Cuentas_Pagar.Mon_Net")  
			loComandoSeleccionar.AppendLine("  				Else @lnCero")				  
			loComandoSeleccionar.AppendLine("  			END												AS Islr,")
			loComandoSeleccionar.AppendLine("  			CASE WHEN Cuentas_Pagar.Cod_Tip = 'RETPAT' ")
			loComandoSeleccionar.AppendLine("  				THEN Cuentas_Pagar.Mon_Net")
			loComandoSeleccionar.AppendLine("  				Else @lnCero")
			loComandoSeleccionar.AppendLine("  			END												AS Patente")
			loComandoSeleccionar.AppendLine("INTO 		#tabRetenciones")
			loComandoSeleccionar.AppendLine("FROM 		Cuentas_Pagar")
			loComandoSeleccionar.AppendLine("JOIN 		Proveedores ")
			loComandoSeleccionar.AppendLine("  		ON	Proveedores.Cod_Pro =  Cuentas_Pagar.Cod_Pro")
			loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip IN ('RETIVA', 'ISLR', 'RETPAT')")
			loComandoSeleccionar.AppendLine(" 		AND	Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)

            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Status <> 'Anulado'")

            loComandoSeleccionar.AppendLine("SELECT		")
            loComandoSeleccionar.AppendLine("  			Cod_Pro										AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("  			Nom_Pro										AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("  			SUM(Impuesto)								AS Impuesto,")
            loComandoSeleccionar.AppendLine("  			SUM(Islr)									AS Islr,")
            loComandoSeleccionar.AppendLine("  			SUM(Patente)								AS Patente,")
            loComandoSeleccionar.AppendLine("  			(SUM(Impuesto) + SUM(Islr) + SUM(Patente))	AS Total")
            loComandoSeleccionar.AppendLine("FROM		#tabRetenciones")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Pro, Nom_Pro")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRetenciones_Proveedores", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrRetenciones_Proveedores.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                 "No se pudo completar el Proceso: " & loExcepcion.Message, _
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
' CMS: 04/08/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' RJG: 03/06/11: Se cambió el filtro de Estatus paraque fuese fijo (diferente a anulado) en '
'				 lugar de ser un parámetro. corregida la interface.							'
'-------------------------------------------------------------------------------------------'
