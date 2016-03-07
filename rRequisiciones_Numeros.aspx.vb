'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRequisiciones_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rRequisiciones_Numeros
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro3Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
		Dim lcParametro4Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro4Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
		Dim lcParametro6Desde	As	String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))			
		Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
		Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
		 Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Requisiciones.Status, ")
            loComandoSeleccionar.appendline("           Requisiciones.Documento, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Pro, ")
            loComandoSeleccionar.appendline("           SUBSTRING(Proveedores.Nom_Pro,1, 30)                        AS Nom_Pro, ")
            loComandoSeleccionar.appendline("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Ven, ")
            loComandoSeleccionar.appendline("           Requisiciones.Mon_Bru, ")
            loComandoSeleccionar.appendline("           (Requisiciones.Mon_Imp1 + Requisiciones.Mon_Imp2 + Requisiciones.Mon_Imp3)    AS Mon_Imp, ")
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
            loComandoSeleccionar.appendline("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.appendline("           AND Proveedores.Cod_Pro             BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Ven           BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Tra            BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Mon            BETWEEN" & lcParametro5Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro5Hasta)
			loComandoSeleccionar.appendline("           AND Requisiciones.Status            IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.appendline("           AND Requisiciones.cod_rev            BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro7Hasta)
          	loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.appendline(" ORDER BY  Requisiciones.Documento ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rRequisiciones_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRequisiciones_Numeros.ReportSource =	 loObjetoReporte	

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
' YJP: 14/05/09: Agregar condición revisión
'-------------------------------------------------------------------------------------------'