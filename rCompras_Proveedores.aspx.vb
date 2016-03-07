Imports System.Data
Partial Class rCompras_Proveedores
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


            loComandoSeleccionar.Append(" SELECT    Compras.Status, ")
            loComandoSeleccionar.Append("           Compras.Documento, ")
            loComandoSeleccionar.Append("           Compras.Cod_Pro, ")
            loComandoSeleccionar.Append("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.Append("           Compras.Fec_Ini, ")
            loComandoSeleccionar.Append("           Compras.Cod_Ven, ")
            loComandoSeleccionar.Append("           Compras.Mon_Bru, ")
            loComandoSeleccionar.Append("           (Compras.Mon_Imp1 + Compras.Mon_Imp2 + Compras.Mon_Imp3) AS Mon_Imp, ")
            loComandoSeleccionar.Append("           Compras.Mon_Net, ")
            loComandoSeleccionar.Append("           Compras.Mon_Sal, ")
            loComandoSeleccionar.Append("           Compras.Cod_Tra, ")
            loComandoSeleccionar.Append("           Transportes.Nom_Tra ")
            loComandoSeleccionar.Append(" FROM      Compras, ")
            loComandoSeleccionar.Append("           Proveedores, ")
            loComandoSeleccionar.Append("           Transportes, ")
            loComandoSeleccionar.Append("           Vendedores ")
            loComandoSeleccionar.Append(" WHERE     Compras.Cod_Pro                     =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.Append("           AND Compras.Cod_Ven                 =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.Append("           AND Compras.Cod_Tra                 =   Transportes.Cod_Tra ")
            loComandoSeleccionar.Append("           AND Compras.Documento               Between " & lcParametro0Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           AND Compras.Fec_Ini                 Between " & lcParametro1Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.Append("           AND Proveedores.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.Append("           AND Compras.Cod_Ven                 Between " & lcParametro3Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.Append("           AND Compras.Cod_Tra                 Between" & lcParametro4Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.Append("           AND Compras.Cod_Mon                 Between" & lcParametro5Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.Append("           AND Compras.Status                  IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Suc                 Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro7Hasta)
            'loComandoSeleccionar.Append(" ORDER BY  Compras.Cod_Pro, Compras.Fec_Ini ")
            loComandoSeleccionar.AppendLine("ORDER BY   Compras.Cod_Pro, " & lcOrdenamiento & ", CONVERT(nchar(30), Compras.Fec_Ini,112)")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCompras_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_Proveedores.ReportSource =	 loObjetoReporte	

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
' GCR: 18/03/09: Estandarización de código y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  20/07/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'