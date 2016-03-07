Imports System.Data
Partial Class rPagos_Proveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
           
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Nom_Pro,1,30) As Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Net   AS  Neto, ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Conceptos.Cod_Con,")
            loComandoSeleccionar.AppendLine("			Conceptos.Nom_Con")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos1 ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos ")
            loComandoSeleccionar.AppendLine(" JOIN      Detalles_Pagos ON (Pagos.Documento	=   Detalles_Pagos.Documento )")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Proveedores ON (Pagos.Cod_Pro       =   Proveedores.Cod_Pro )")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Conceptos ON (Proveedores.Cod_Con   =   Conceptos.Cod_Con )")
            loComandoSeleccionar.AppendLine(" WHERE     ")
            loComandoSeleccionar.AppendLine("			Pagos.Documento	        Between	" & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Fec_Ini	        Between	" & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Pro           Between	" & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Ven           Between	" & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Mon           Between	" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Status	        IN	(" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Suc           Between	" & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Proveedores.Cod_Con     Between	" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine(" ORDER BY  Pagos.Cod_Pro, Pagos.Fec_Ini, Pagos.Documento, Detalles_Pagos.Tip_Ope, Detalles_Pagos.Renglon")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj   AS  Nom_Caj ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban   AS  Nom_Ban ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos3 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")


            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue   AS  Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos4 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loComandoSeleccionar.AppendLine("           Tarjetas.Nom_Tar   AS  Nom_Tar ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")
            loComandoSeleccionar.AppendLine("ORDER BY    Cod_Pro, " & lcOrdenamiento)

	
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 22/04/09: Corregir estatus, anexar combo
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' MAT:  17/03/11: Filtro "Concepto:", Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de registros.											'
'-------------------------------------------------------------------------------------------'
