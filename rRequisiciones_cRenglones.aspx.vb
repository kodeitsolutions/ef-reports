'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRequisiciones_cRenglones"
'-------------------------------------------------------------------------------------------'
Partial Class rRequisiciones_cRenglones
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden



	Try

		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde	As	String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))			
		Dim lcParametro4Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro4Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Requisiciones.Documento, ")
            loComandoSeleccionar.appendline("         Requisiciones.status, ")
            loComandoSeleccionar.appendline("           Requisiciones.Cod_Pro, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Renglon, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Mon_Bru, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (Renglones_Requisiciones.Mon_Net + Renglones_Requisiciones.Mon_Imp1) AS Mon_Net, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Mon_Sal, ")
            loComandoSeleccionar.appendline("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Cod_Art, ")
            loComandoSeleccionar.appendline("           SUBSTRING(Articulos.Nom_Art,1,35) AS Nom_Art, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones.Can_Art1 ")
            loComandoSeleccionar.appendline(" FROM      Requisiciones, ")
            loComandoSeleccionar.appendline("           Renglones_Requisiciones, ")
            loComandoSeleccionar.appendline("           Proveedores, ")
            loComandoSeleccionar.appendline("           Articulos ")
            loComandoSeleccionar.appendline(" WHERE     Requisiciones.Documento             =   Renglones_Requisiciones.Documento ")
            loComandoSeleccionar.appendline("           AND Renglones_Requisiciones.Cod_Art =   Articulos.Cod_Art ")
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Pro           =   Proveedores.Cod_Pro " )
            loComandoSeleccionar.appendline("           AND Renglones_Requisiciones.Cod_Art =   Articulos.Cod_Art ")
            loComandoSeleccionar.appendline("           AND Requisiciones.Documento         BETWEEN " & lcParametro0Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro0Hasta )
            loComandoSeleccionar.appendline("           AND Requisiciones.Fec_Ini           BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_Pro           BETWEEN " & lcParametro2Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro2Hasta )
            loComandoSeleccionar.appendline("           AND Requisiciones.Status            IN ( " & lcParametro3Desde & ")")
            loComandoSeleccionar.appendline("           AND Requisiciones.Cod_rev           BETWEEN " & lcParametro4Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro4Hasta )
            'loComandoSeleccionar.appendline(" ORDER BY  Renglones_Requisiciones.Documento, Renglones_Requisiciones.Renglon ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            Dim loServicios As New cusDatos.goDatos

            'Me.mescribirconsulta(loComandoSeleccionar.toString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rRequisiciones_cRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRequisiciones_cRenglones.ReportSource =	 loObjetoReporte	

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
' CMS: 22/05/09: Se Ajusto el Monto Neto
'                (Renglones_Requisiciones.Mon_Net + Renglones_Requisiciones.Mon_Imp1) AS Mon_Net
'-------------------------------------------------------------------------------------------'
' AAP: 24/06/09: metodo de ordenamiento
'-------------------------------------------------------------------------------------------'

