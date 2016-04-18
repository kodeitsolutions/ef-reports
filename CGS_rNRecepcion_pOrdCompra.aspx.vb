Imports System.Data
Partial Class CGS_rNRecepcion_pOrdCompra

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("SELECT	DISTINCT	Recepciones.Documento		AS Recepcion,")
            lcComandoSeleccionar.AppendLine("			Recepciones.Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, ")
            lcComandoSeleccionar.AppendLine("			Recepciones.Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("			RR.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("			(SELECT SUM(Can_Art1)")
            lcComandoSeleccionar.AppendLine("			FROM Renglones_Recepciones")
            lcComandoSeleccionar.AppendLine("			WHERE Cod_Art = RR.Cod_Art")
            lcComandoSeleccionar.AppendLine("			AND Documento = RR.Documento)		AS Cantidad_Recibida,")
            lcComandoSeleccionar.AppendLine("			(SELECT SUM(Can_Art1)")
            lcComandoSeleccionar.AppendLine("			FROM Renglones_OCompras")
            lcComandoSeleccionar.AppendLine("			WHERE Cod_Art = ROC.Cod_Art")
            lcComandoSeleccionar.AppendLine("			AND Documento = ROC.Documento)		AS Cantidad_Ordenada,")
            lcComandoSeleccionar.AppendLine("			RR.Cod_Uni							AS Unidad,")
            lcComandoSeleccionar.AppendLine("			ROC.Documento						AS Orden_Compra,")
            'lcComandoSeleccionar.AppendLine("			CASE WHEN (Renglones_Recepciones.Can_Art1 - Renglones_OCompras.Can_Art1) = 0")
            'lcComandoSeleccionar.AppendLine("			THEN 'TOTAL' ELSE 'PARCIAL' END		AS Entrega,")
            lcComandoSeleccionar.AppendLine("			Operaciones_Lotes.Cod_Lot			AS Lote,")
            lcComandoSeleccionar.AppendLine("			Operaciones_Lotes.Cantidad			AS Cantidad_Lote")
            lcComandoSeleccionar.AppendLine("FROM	 Recepciones ")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recepciones AS RR ON Recepciones.Documento	= RR.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Recepciones.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON RR.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine("   JOIN Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep")
            lcComandoSeleccionar.AppendLine("   JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            lcComandoSeleccionar.AppendLine("   	AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_OCompras AS ROC ON ROC.Documento =  RR.Doc_Ori")
            lcComandoSeleccionar.AppendLine("		AND ROC.Renglon = RR.Ren_Ori")
            lcComandoSeleccionar.AppendLine("	JOIN Ordenes_Compras ON ROC.Documento = Ordenes_Compras.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = RR.Documento")
            lcComandoSeleccionar.AppendLine("		AND RR.Cod_Art = Operaciones_Lotes.Cod_Art")
            lcComandoSeleccionar.AppendLine("WHERE	Recepciones.Fec_Ini	Between	" & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("	        AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Pro	Between	" & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("		AND RR.Cod_Art Between	" & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("		AND Departamentos.Cod_Dep Between	" & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("		AND Secciones.Cod_Sec Between	" & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("		AND Recepciones.Documento Between	" & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("		AND Ordenes_Compras.Documento Between	" & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("		    AND " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rNRecepcion_pOrdCompra", laDatosReporte)

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
