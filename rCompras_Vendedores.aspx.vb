Imports System.Data
Partial Class rCompras_Vendedores
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
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
		

            loComandoSeleccionar.AppendLine(" SELECT    Compras.Status, ")
            loComandoSeleccionar.AppendLine("           Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Compras.Mon_Imp1 + Compras.Mon_Imp2 + Compras.Mon_Imp3)    AS Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Sal, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra ")
            loComandoSeleccionar.AppendLine(" FROM      Compras, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Compras.Cod_Pro                     =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven                 =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Tra                 =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Compras.Documento               Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini                 Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven                 Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Tra                 Between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Mon                 Between" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Status                  IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_rev                 Between" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Compras.Cod_Ven, Compras.Fec_Ini, Compras.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    Compras.Cod_Ven, " & lcOrdenamiento)


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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCompras_Vendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_Vendedores.ReportSource =	 loObjetoReporte	

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
' JJD: 09/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 17/02/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 18/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'