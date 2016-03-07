Imports System.Data
Partial Class rPresupuestos_Fechas
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
		Dim lcParametro4Desde	As  String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
		Dim lcParametro6Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
		Dim lcParametro6Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
		Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' 1 - Select Principal
            '-------------------------------------------------------------------------------------------'
			loComandoSeleccionar.AppendLine(" SELECT	Presupuestos.Documento, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Status, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Tra, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Mon, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Control, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Presupuestos.Mon_Sal  ")
			loComandoSeleccionar.AppendLine(" FROM		Proveedores, ")
			loComandoSeleccionar.AppendLine("			Presupuestos, ")
			loComandoSeleccionar.AppendLine("			Vendedores, ")
			loComandoSeleccionar.AppendLine("			Transportes, ")
			loComandoSeleccionar.AppendLine("			Monedas ")
			loComandoSeleccionar.AppendLine(" WHERE		Presupuestos.Cod_Pro		=	Proveedores.Cod_Pro ")
			loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Ven	=	Vendedores.Cod_Ven ")
			loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Tra	=	Transportes.Cod_Tra ")
			loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Mon	=	Monedas.Cod_Mon ")
			loComandoSeleccionar.AppendLine("			And Presupuestos.Documento	Between " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			And Presupuestos.Fec_Ini	Between " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			And Proveedores.Cod_Pro		Between " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("			And Vendedores.Cod_Ven		Between " & lcParametro3Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			And Presupuestos.Status		IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine("			And Transportes.Cod_Tra		Between " & lcParametro5Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("			And Monedas.Cod_Mon			Between " & lcParametro6Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta )
			loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_rev			Between " & lcParametro7Desde )
			loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta )
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			'loComandoSeleccionar.AppendLine(" ORDER BY  Presupuestos.Documento ")
			


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPresupuestos_Fechas", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPresupuestos_Fechas.ReportSource = loObjetoReporte

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
' JJD: 01/11/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 28/04/09: Agregar combo estatus
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revision
'-------------------------------------------------------------------------------------------'
' MAT:  04/04/11: Mejora de la vista de diseño.
'-------------------------------------------------------------------------------------------'