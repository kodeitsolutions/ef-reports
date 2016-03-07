'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTAlmacenes_Costos"
'-------------------------------------------------------------------------------------------'
Partial Class rTAlmacenes_Costos
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = cusAplicacion.goReportes.paParametrosIniciales(4)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            Dim lcCosto As String = ""

            Select Case lcParametro4Desde
                Case "Promedio MP"
                    lcCosto = "Cos_Pro1"
                Case "Ultimo MP"
                    lcCosto = "Cos_Ult1"
                Case "Anterior MP"
                    lcCosto = "Cos_Ant1"
                Case "Promedio MS"
                    lcCosto = "Cos_Pro2"
                Case "Ultimo MS"
                    lcCosto = "Cos_Ult2"
                Case "Anterior MS"
                    lcCosto = "Cos_Ant2"
            End Select

            loComandoSeleccionar.AppendLine("SELECT	MAX(Auditorias.Registro) AS Fecha_Confirmacion, ")
            loComandoSeleccionar.AppendLine("		Auditorias.Documento ")
            loComandoSeleccionar.AppendLine("INTO #tmpTemporal ")
            loComandoSeleccionar.AppendLine("FROM Auditorias ")
            loComandoSeleccionar.AppendLine("WHERE 	Auditorias.Tabla = 'Traslados' ")
            loComandoSeleccionar.AppendLine("		AND Auditorias.Accion = 'Confirmar' ")
            loComandoSeleccionar.AppendLine("GROUP BY Auditorias.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY Documento ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine(" SELECT	Traslados.Documento, ")
            loComandoSeleccionar.AppendLine("			Traslados.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Traslados.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Traslados.Alm_Ori, ")
            loComandoSeleccionar.AppendLine("			Traslados.Alm_Des, ")
            loComandoSeleccionar.AppendLine("			Traslados.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Traslados.Tasa, ")
            loComandoSeleccionar.AppendLine("			Traslados.Status, ")
            loComandoSeleccionar.AppendLine("			Traslados.Comentario, ")
            loComandoSeleccionar.AppendLine("			Traslados.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("			Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Traslados." & lcCosto & " AS Costo,")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("		    ISNULL(#tmpTemporal.Fecha_Confirmacion,0) AS Fecha_Confirmacion, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine(" FROM	Traslados ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Traslados ON Traslados.Documento	=	Renglones_Traslados.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Traslados.Cod_Art		=	Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" JOIN Transportes ON Traslados.Cod_Tra	=	Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine(" LEFT JOIN #tmpTemporal ON (#tmpTemporal.Documento = Traslados.Documento)")
            loComandoSeleccionar.AppendLine(" WHERE		Traslados.Documento				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Traslados.Fec_Ini				BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Traslados.Status				IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Traslados.Alm_Ori				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Traslados.Alm_Des				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("      	    AND Traslados.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal")

            Dim loServicios As New cusDatos.goDatos

            ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTAlmacenes_Costos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTAlmacenes_Costos.ReportSource = loObjetoReporte

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
' MAT: 14/09/11: Código inicial
'-------------------------------------------------------------------------------------------'
