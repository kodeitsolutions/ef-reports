'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Vendidos1"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Vendidos1

    Inherits vis2Formularios.frmReporte


    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("WITH curTemporal AS ( ")
        loComandoSeleccionar.AppendLine("SELECT		")
        loComandoSeleccionar.AppendLine("			Facturas.Fec_Ini, ")
        loComandoSeleccionar.AppendLine("			Renglones_Facturas.Cod_Art, ")
        loComandoSeleccionar.AppendLine("			Renglones_Facturas.Cod_Alm, ")
        loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
        loComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
        loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar, ")
        loComandoSeleccionar.AppendLine("			(Renglones_Facturas.Can_Art1) As Can_Art, ")
        loComandoSeleccionar.AppendLine("			(Renglones_Facturas.Mon_Net)  As Mon_Net ")
        loComandoSeleccionar.AppendLine("FROM		Facturas, Renglones_Facturas, Articulos ")
        loComandoSeleccionar.AppendLine("WHERE		Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
        loComandoSeleccionar.AppendLine("			AND Facturas.Documento = Renglones_Facturas.Documento and Renglones_Facturas.Cod_Art = Articulos.Cod_Art ")
        loComandoSeleccionar.AppendLine("			AND Facturas.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia))
        loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Art BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1)))
        loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1)))
        loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Alm BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)))
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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Vendidos1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Vendidos1.ReportSource = loObjetoReporte


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
' JFP: 05/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP: 26/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 27/04/10: Se Aplico el redondeo de los parametros Fechas 
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro.					' 
'-------------------------------------------------------------------------------------------'
' MAT: 14/02/11: Ajustado el select y el criterio de ordenamiento		.					' 
'-------------------------------------------------------------------------------------------'
