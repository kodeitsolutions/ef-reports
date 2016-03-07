'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rComisiones_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rComisiones_Vendedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Facturas.Cod_Ven,	")
            loComandoSeleccionar.AppendLine("       	Vendedores.Nom_Ven,	")
            loComandoSeleccionar.AppendLine("       	Vendedores.Por_Ven,	")
            loComandoSeleccionar.AppendLine("       	COUNT(Facturas.Cod_Ven)									AS	Cuantos, ")
            loComandoSeleccionar.AppendLine("       	SUM(Facturas.Mon_Bru)                             		AS	Mon_Bru, ")
            loComandoSeleccionar.AppendLine("       	SUM(Facturas.Mon_Net)                             		AS	Mon_Net, ")
            loComandoSeleccionar.AppendLine("       	SUM(Facturas.Mon_Bru * (Vendedores.Por_Ven / 100))		AS	Mon_Com, ")
            loComandoSeleccionar.AppendLine("       	SUM(Renglones_Facturas.Can_Art1)                  		AS	Pares ")
            loComandoSeleccionar.AppendLine("FROM		Facturas,			")
            loComandoSeleccionar.AppendLine("       	Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("       	Vendedores			")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.status			IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("	AND		Facturas.Documento      =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("	AND		Facturas.Cod_Ven		=   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("	AND		Facturas.Fec_Ini	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("	AND		Facturas.Cod_Ven    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Facturas.Cod_Ven, Vendedores.Nom_Ven, Vendedores.Por_Ven ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComisiones_Vendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComisiones_Vendedores.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 28/02/09: Normalizacion del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 14/04/09: Código de Ordenamiento, Verificación de registros
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro.					' 
'-------------------------------------------------------------------------------------------'
