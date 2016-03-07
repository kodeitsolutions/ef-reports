'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRequisiciones_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rRequisiciones_Proveedores
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try
	
		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro3Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro3Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro4Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro4Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro6Desde	As	String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))		
	
		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Requisiciones.Status, ")
            loComandoSeleccionar.appendline("           Requisiciones.Documento, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Pro, ")
            loComandoSeleccionar.appendline("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Ven, ")
            loComandoSeleccionar.appendline("           Requisiciones.Mon_Bru, ")
            loComandoSeleccionar.appendline("           (Requisiciones.Mon_Imp1 + Requisiciones.Mon_Imp2 + Requisiciones.Mon_Imp3) AS Mon_Imp, ")
            loComandoSeleccionar.appendline("           Requisiciones.Mon_Net, ")
            loComandoSeleccionar.appendline("           Requisiciones.Mon_Sal, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Tra, ")
            loComandoSeleccionar.appendline("           Transportes.Nom_Tra ")
            loComandoSeleccionar.appendline(" FROM      Requisiciones, ")
            loComandoSeleccionar.appendline("           Proveedores, ")
            loComandoSeleccionar.appendline("           Transportes, ")
            loComandoSeleccionar.appendline("           Vendedores ")
            loComandoSeleccionar.appendline(" WHERE     Requisiciones.Cod_Pro               =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Ven           =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Tra           =   Transportes.Cod_Tra ")
            loComandoSeleccionar.appendline("           AND Requisiciones.Documento         BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Fec_Ini           BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro1Hasta )
            loComandoSeleccionar.appendline("           AND Proveedores.Cod_Pro             BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Ven           BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Tra           BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Mon           BETWEEN" & lcParametro5Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Requisiciones.Status            IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.appendline(" ORDER BY  Requisiciones.Cod_Pro, Requisiciones.Fec_Ini ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rRequisiciones_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRequisiciones_Proveedores.ReportSource =	 loObjetoReporte	

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
' AAP: 24/06/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'