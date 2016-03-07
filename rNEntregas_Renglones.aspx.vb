'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNEntregas_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rNEntregas_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))            
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Entregas.Documento, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,20)      AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,26)     AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine(" FROM      Entregas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Entregas.Documento      =   Renglones_Entregas.Documento ")
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art   =   Renglones_Entregas.Cod_Art ")
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Tra    =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Entregas.Documento  BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Entregas.Fec_Ini    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cli    BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
			loComandoSeleccionar.Appendline("			AND Vendedores.Cod_Ven between " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("			AND " & lcParametro3Hasta )
            loComandoSeleccionar.AppendLine("           AND Entregas.Status     IN ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Rev between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Suc between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro6Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Entregas.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY   Entregas.Documento, " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNEntregas_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNEntregas_Renglones.ReportSource = loObjetoReporte

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
' JJD: 19/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/04/09: Cambios Estandarización de codigo. 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' CMS: 27/04/10: se agrego el filtro Vendedores
'-------------------------------------------------------------------------------------------'