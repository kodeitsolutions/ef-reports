Imports System.Data
Partial Class rPagos_Fechas

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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosFinales(9)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT		Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("			Pagos.Status, ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Proveedores.Nom_Pro,1,40) As Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Mon_Net   AS  Neto, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Sal, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Renglon, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Tip_Ope, ")
            
            loComandoSeleccionar.AppendLine("			ISNULL((SELECT nom_caj FROM cajas WHERE Cajas.Cod_Caj = Detalles_Pagos.Cod_Caj),'') AS Nom_Caj, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Num_Doc,  ")
            loComandoSeleccionar.AppendLine("			ISNULL((SELECT Num_Cue FROM Cuentas_Bancarias WHERE Cuentas_Bancarias.Cod_Cue = Detalles_Pagos.Cod_Cue),'') As Num_Cue,  ")
            
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("			Detalles_Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			CASE WHEN Detalles_Pagos.Tip_Ope = 'Efectivo' THEN 'Caja' ELSE 'Banco' END AS Tipo_Renglon ")
            loComandoSeleccionar.AppendLine("FROM		Pagos")
            loComandoSeleccionar.AppendLine("	JOIN	Detalles_Pagos ON Pagos.Documento = Detalles_Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Pagos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("WHERE			Pagos.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Pro       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Mon       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Status        IN (" & lcParametro4Desde & ")")
            
            
            loComandoSeleccionar.AppendLine("           AND Detalles_Pagos.Cod_Caj    BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Detalles_Pagos.Cod_Cue    BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc    BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            
            If lcParametro9Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("           AND Pagos.Cod_rev   BETWEEN " & lcParametro8Desde)
            Else
                loComandoSeleccionar.AppendLine("           AND Pagos.Cod_rev  NOT BETWEEN " & lcParametro8Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
																												   
            loComandoSeleccionar.AppendLine("ORDER BY    Pagos.Fec_Ini, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Fechas.ReportSource = loObjetoReporte

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
' JJD: 09/10/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' GCR: 30/03/09: Estandarizacion de codigo, adicion del combo estatus y ajustes al diseño	'
'-------------------------------------------------------------------------------------------'
' AAP: 01/07/09: Filtro "Sucursal:"															'
'-------------------------------------------------------------------------------------------'
' CMS: 10/08/09: Metodo de ordenamiento, verificacionde registros							'
'-------------------------------------------------------------------------------------------'
' CMS: 03/07/10: Se agregaron los campos numero de cuenta y caja							'
'-------------------------------------------------------------------------------------------'
' RJG: 10/04/12: Se agregó el total de documentos.											'
'-------------------------------------------------------------------------------------------'
