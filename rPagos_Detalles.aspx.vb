'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPagos_Detalles"
'-------------------------------------------------------------------------------------------'
Partial Class rPagos_Detalles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosFinales(8)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		Pagos.Cod_Pro											AS Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro										AS Nom_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Rif											AS Rif, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nit											AS Nit, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis										AS Dir_Fis, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Telefonos									AS Telefonos, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Fax											AS Fax, ")
			loComandoSeleccionar.AppendLine("			Pagos.Documento											AS Documento, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini											AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin											AS Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Bru											AS Mon_Bru, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Imp											AS Mon_Imp, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Net											AS Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Pagos.Comentario										AS Comentario,")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Renglon      							AS Ren_Tip, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Tip_Ope      							AS Tip_Ope, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Doc_Des      							AS Doc_Des, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Num_Doc      							AS Num_Doc, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Caj      							AS Cod_Caj, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Ban      							AS Cod_Ban, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Cue      							AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Tar      							AS Cod_Tar, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Mon_Net      							AS Mon_Net_Det,")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Cajas.Nom_Caj,1,25),'')				AS Nom_Caj,")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Bancos.Nom_Ban,1,25),'')				AS Nom_Ban,")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25),'')	AS Nom_Cue,")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Tarjetas.Nom_Tar,1,25),'')				AS Nom_Tar ")
			loComandoSeleccionar.AppendLine("FROM		Pagos ")
			loComandoSeleccionar.AppendLine("	JOIN	Proveedores ")
			loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Pagos.Cod_Pro")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Detalles_Pagos")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Documento = Pagos.Documento")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Cajas ")
			loComandoSeleccionar.AppendLine("		ON  Detalles_Pagos.Cod_Caj = Cajas.Cod_Caj ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Cod_Ban = Bancos.Cod_Ban ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Bancarias ")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Cod_Cue = Cuentas_Bancarias.Cod_Cue ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Tarjetas ")
			loComandoSeleccionar.AppendLine("		ON  Detalles_Pagos.Cod_Tar = Tarjetas.Cod_Tar ")
			loComandoSeleccionar.AppendLine("WHERE		Pagos.Documento			BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Fec_Ini			BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Pro			BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Status			IN (" & lcParametro3Desde & ")")
			loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Suc			BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			
			If (lcParametro5Desde <> "''") OrElse (lcParametro5Hasta <> "'zzzzzzz'") Then 
				loComandoSeleccionar.AppendLine("		AND Detalles_Pagos.Cod_Caj	BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			End If
			
			If (lcParametro6Desde <> "''") OrElse (lcParametro6Hasta <> "'zzzzzzz'") Then 
 				loComandoSeleccionar.AppendLine("		AND Detalles_Pagos.Cod_Cue	BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
			End If
  
			If lcParametro8Desde = "Igual" Then
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			Else
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev 			NOT BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			End If
          
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			    
            

            Dim loServicios As New cusDatos.goDatos
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_Detalles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Detalles.ReportSource = loObjetoReporte

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
' RJG: 27/04/12: Programacion inicial, a partir de rPagos_Detalles.							'
'-------------------------------------------------------------------------------------------'
 