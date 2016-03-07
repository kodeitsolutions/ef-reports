'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Clientes1"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Clientes1
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Facturas.Cod_Cli AS	Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli AS	Nom_Cli,")
            loComandoSeleccionar.AppendLine("	    Renglones_Facturas.Can_Art1 AS Can_Art,")
            loComandoSeleccionar.AppendLine("	    Renglones_Facturas.Mon_Net AS Mon_Net")
            loComandoSeleccionar.AppendLine("INTO	#tabla_ART_MON")
            loComandoSeleccionar.AppendLine("FROM")
            loComandoSeleccionar.AppendLine("		Facturas,")
            loComandoSeleccionar.AppendLine("		Renglones_Facturas,")
            loComandoSeleccionar.AppendLine("		Articulos,")
            loComandoSeleccionar.AppendLine("		Vendedores,")
            loComandoSeleccionar.AppendLine("		Clientes")
            loComandoSeleccionar.AppendLine("WHERE ")
            loComandoSeleccionar.AppendLine("		Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("        AND Facturas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("        AND Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("        AND Facturas.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("        AND Facturas.Fec_Ini BETWEEN  " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("        AND Facturas.Cod_Cli BETWEEN  " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("        AND Facturas.Cod_Ven BETWEEN  " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("        AND Articulos.Cod_Art BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("        AND Facturas.Cod_Rev BETWEEN  " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("        AND Facturas.Cod_Suc BETWEEN  " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("		Facturas.Cod_Cli AS	Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Facturas.Documento AS Documento")
            loComandoSeleccionar.AppendLine("INTO	#tablaDOCUMENTOS")
            loComandoSeleccionar.AppendLine("FROM")
            loComandoSeleccionar.AppendLine("		Facturas,")
            loComandoSeleccionar.AppendLine("		Renglones_Facturas,")
            loComandoSeleccionar.AppendLine("		Articulos,")
            loComandoSeleccionar.AppendLine("		Vendedores,")
            loComandoSeleccionar.AppendLine("		Clientes")
            loComandoSeleccionar.AppendLine("WHERE ")
            loComandoSeleccionar.AppendLine("		Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Facturas.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini BETWEEN  " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli BETWEEN  " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven BETWEEN  " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Art BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Rev BETWEEN  " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc BETWEEN  " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Cod_Cli AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		COUNT(Documento) AS Documento")
            loComandoSeleccionar.AppendLine("INTO	#tablaCOUNT_DOC")
            loComandoSeleccionar.AppendLine("FROM	#tablaDOCUMENTOS")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine("		#tabla_ART_MON.Cod_Cli AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		#tabla_ART_MON.Nom_Cli AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		#tablaCOUNT_DOC.Documento AS Documentos,")
            loComandoSeleccionar.AppendLine("		SUM(#tabla_ART_MON.Can_Art) AS Can_Art, ")
            loComandoSeleccionar.AppendLine("		SUM(#tabla_ART_MON.Mon_Net) AS Mon_Net,")
            loComandoSeleccionar.AppendLine(" 	    " & goOpciones.pnDecimalesParaCantidad() & " AS NumDecCant")
            loComandoSeleccionar.AppendLine("FROM	#tabla_ART_MON")
            loComandoSeleccionar.AppendLine("JOIN	#tablaCOUNT_DOC ON #tabla_ART_MON.Cod_Cli = #tablaCOUNT_DOC.Cod_Cli")
            loComandoSeleccionar.AppendLine("GROUP BY #tabla_ART_MON.Cod_Cli, #tabla_ART_MON.Nom_Cli, #tablaCOUNT_DOC.Documento")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Clientes1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Clientes1.ReportSource = loObjetoReporte

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
' DLC: 30/06/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 20/07/2010: Ajuste de la consulta a la base de datos, acomodando el conteno de los 
'                   documentos.
'-------------------------------------------------------------------------------------------'
' MAT: 16/02/11: Rediseño de la vista del reporte. 
'-------------------------------------------------------------------------------------------'