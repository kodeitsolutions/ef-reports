'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Comprados"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Comprados

    Inherits vis2Formularios.frmReporte


    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("WITH curTemporal AS ( ")
        loComandoSeleccionar.AppendLine("SELECT		")
        loComandoSeleccionar.AppendLine("			Compras.Fec_Ini, ")
        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art, ")
        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm, ")
        loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
        loComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
        loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar, ")
        loComandoSeleccionar.AppendLine("			(Renglones_Compras.Can_Art1) As Can_Art, ")
        loComandoSeleccionar.AppendLine("			(Renglones_Compras.Mon_Net)  As Mon_Net ")
        loComandoSeleccionar.AppendLine("FROM		Compras, Renglones_Compras, Articulos ")
        loComandoSeleccionar.AppendLine("WHERE		Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
        loComandoSeleccionar.AppendLine("			AND Compras.Documento = Renglones_Compras.Documento and Renglones_Compras.Cod_Art = Articulos.Cod_Art ")
        loComandoSeleccionar.AppendLine("			AND Compras.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia))
        loComandoSeleccionar.AppendLine("			AND Renglones_Compras.Cod_Art BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1)))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1)))
        loComandoSeleccionar.AppendLine("			AND Renglones_Compras.Cod_Alm BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2)))
        loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3)))
        loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4)))
        loComandoSeleccionar.AppendLine(") ")
							  
        loComandoSeleccionar.AppendLine("SELECT			")
        loComandoSeleccionar.AppendLine("			Cod_Art,")
        loComandoSeleccionar.AppendLine("			Nom_Art,")
        loComandoSeleccionar.AppendLine("			SUM(Can_Art)	AS	Can_Art, ")
        loComandoSeleccionar.AppendLine("			SUM(Mon_Net)	AS  Mon_Net ")
        loComandoSeleccionar.AppendLine("FROM		curTemporal ")
        loComandoSeleccionar.AppendLine("GROUP BY	Cod_Art, Nom_Art ")
        loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)


        Try


            Dim loServicios As New cusDatos.goDatos

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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Comprados", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Comprados.ReportSource = loObjetoReporte


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
' MAT: 01/04/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
