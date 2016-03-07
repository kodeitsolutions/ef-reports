Imports System.Data
Partial Class rNRecepcion_cRenglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine(" SELECT	Recepciones.Documento, ")
            lcComandoSeleccionar.AppendLine("			Recepciones.Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Renglon, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Mon_Net, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Mon_Sal, ")
            lcComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, ")
            lcComandoSeleccionar.AppendLine("			Recepciones.Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones.Can_Art1 ")
            lcComandoSeleccionar.AppendLine(" FROM		Recepciones, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Recepciones, ")
            lcComandoSeleccionar.AppendLine("			Proveedores, ")
            lcComandoSeleccionar.AppendLine("			Articulos ")
            lcComandoSeleccionar.AppendLine(" WHERE		Recepciones.Documento		        =	Renglones_Recepciones.Documento ")
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Cod_Pro		        =	Proveedores.Cod_Pro ")
            lcComandoSeleccionar.AppendLine("			AND Renglones_Recepciones.Cod_Art	=	Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Documento	        Between	" & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Fec_Ini				Between	" & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Cod_Pro				Between	" & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Status				IN	(" & lcParametro3Desde & ")")
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Cod_rev				Between	" & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("			AND Recepciones.Cod_Suc				Between	" & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            'lcComandoSeleccionar.AppendLine(" ORDER BY	Recepciones.Documento ")
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNRecepcion_cRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNRecepcion_cRenglones.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
