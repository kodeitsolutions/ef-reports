'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRequisiciones_Fechas"
'-------------------------------------------------------------------------------------------'

Partial Class rRequisiciones_Fechas
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
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


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append(" SELECT    Requisiciones.Status, ")
            loComandoSeleccionar.Append("           Requisiciones.Documento, ")
            loComandoSeleccionar.Append("           Requisiciones.Cod_Pro, ")
            loComandoSeleccionar.Append("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.Append("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.Append("           Requisiciones.Cod_Ven, ")
            loComandoSeleccionar.Append("           Requisiciones.Mon_Bru, ")
            loComandoSeleccionar.Append("           (Requisiciones.Mon_Imp1 + Requisiciones.Mon_Imp2 + Requisiciones.Mon_Imp3) AS Mon_Imp, ")
            loComandoSeleccionar.Append("           Requisiciones.Mon_Net, ")
            loComandoSeleccionar.Append("           Requisiciones.Mon_Sal, ")
            loComandoSeleccionar.Append("           Requisiciones.Cod_Tra, ")
            loComandoSeleccionar.Append("           Transportes.Nom_Tra ")
            loComandoSeleccionar.Append(" FROM      Requisiciones, ")
            loComandoSeleccionar.Append("           Proveedores, ")
            loComandoSeleccionar.Append("           Transportes, ")
            loComandoSeleccionar.Append("           Vendedores ")
            loComandoSeleccionar.Append(" WHERE     Requisiciones.Cod_Pro               =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Ven           =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Tra           =   Transportes.Cod_Tra ")
            loComandoSeleccionar.Append("           AND Requisiciones.Documento         Between " & lcParametro0Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Fec_Ini           Between " & lcParametro1Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.Append("           AND Proveedores.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Ven           Between " & lcParametro3Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Tra           Between" & lcParametro4Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Mon           Between" & lcParametro5Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Status            IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_rev           Between" & lcParametro7Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro7Hasta)
            'loComandoSeleccionar.Append(" ORDER BY  Requisiciones.Fec_Ini, Requisiciones.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rRequisiciones_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRequisiciones_Fechas.ReportSource =	 loObjetoReporte	

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
' GCR: 16/09/08: Estandarizacion de codigo   y ajustes al diseño.
'-------------------------------------------------------------------------------------------'
' AAP: 24/06/08: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
