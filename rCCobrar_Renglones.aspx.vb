'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_Renglones

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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Status, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Renglon                AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Art                AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Precio1                AS  Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Can_Art                AS  Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Uni                AS  Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Mon_Net                AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Comentario             AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,60)          AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,100)        AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)        AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Tip                AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Tipos_Documentos.Nom_Tip,1,40)    AS  Nom_Tip ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Documento        =   Renglones_Documentos.Documento ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      =   Tipos_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.cod_tip        =   Renglones_Documentos.cod_tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_For      =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_Documentos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Renglones_Documentos.Origen =   'Ventas' ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Ven      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Zon            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Status       IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Tip            Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cla            Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tra      Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Mon      Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("		    AND ((" & lcParametro11Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("			OR (" & lcParametro11Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Rev between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Cobrar.Cod_Tip, Renglones_Documentos.Documento," & lcOrdenamiento)

            'loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Cobrar.Cod_Tip, Renglones_Documentos.Documento, Renglones_Documentos.Renglon")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_Renglones.ReportSource = loObjetoReporte

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
' JJD: 07/03/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' CMS: 20/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'
