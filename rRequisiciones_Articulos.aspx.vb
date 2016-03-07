Imports System.Data
Partial Class rRequisiciones_Articulos
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
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
			Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
			Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))



            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.Append("           Articulos.Nom_Art, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Sec, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.Append("           Requisiciones.Status, ")
            loComandoSeleccionar.Append("           Requisiciones.Documento, ")
            loComandoSeleccionar.Append("           Requisiciones.Cod_Pro, ")
            loComandoSeleccionar.Append("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.Append("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.Append("           Requisiciones.Cod_Ven, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Renglon, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Cod_Alm, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Can_Art1, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Cod_Uni, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Precio1, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Por_Des, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones.Mon_Net ")
            loComandoSeleccionar.Append(" FROM      Articulos, ")
            loComandoSeleccionar.Append("           Requisiciones, ")
            loComandoSeleccionar.Append("           Renglones_Requisiciones, ")
            loComandoSeleccionar.Append("           Proveedores, ")
            loComandoSeleccionar.Append("           Vendedores, ")
            loComandoSeleccionar.Append("           Almacenes, ")
            loComandoSeleccionar.Append("           Departamentos, ")
            loComandoSeleccionar.Append("           Secciones, ")
            loComandoSeleccionar.Append("           Marcas ")
            loComandoSeleccionar.Append(" WHERE     Requisiciones.Documento             =   Renglones_Requisiciones.Documento ")
            loComandoSeleccionar.Append("           AND Renglones_Requisiciones.Cod_Art =   Articulos.Cod_Art ")
            loComandoSeleccionar.Append("           AND Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.Append("           AND Articulos.Cod_Sec               =   Secciones.Cod_Sec ")
            loComandoSeleccionar.Append("           AND Secciones.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Pro           =   Proveedores.Cod_Pro " )
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Ven           =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.Append("           AND Renglones_Requisiciones.Cod_Alm =   Almacenes.Cod_Alm ")
            loComandoSeleccionar.Append("           AND Renglones_Requisiciones.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Fec_Ini       	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.Append("           AND Proveedores.Cod_Pro         	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Cod_Ven       	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.Append("           AND Articulos.Cod_Dep           	BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.Append("           AND Articulos.Cod_Sec           	BETWEEN" & lcParametro5Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.Append("           AND Articulos.Cod_Mar           	BETWEEN" & lcParametro6Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.Append("           AND Requisiciones.Status        	IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.Append("           AND Renglones_Requisiciones.Cod_Alm BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro8Hasta)
             loComandoSeleccionar.Append("           AND Requisiciones.Cod_rev BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro9Hasta)
            'loComandoSeleccionar.Append(" ORDER BY  Renglones_Requisiciones.Cod_Art, Requisiciones.Fec_Ini, Requisiciones.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rRequisiciones_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRequisiciones_Articulos.ReportSource =	 loObjetoReporte	

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
' AAP: 24/06/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'