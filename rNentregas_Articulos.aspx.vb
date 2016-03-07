'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNentregas_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rNentregas_Articulos
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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT  Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("		Articulos.Status, ")
            loComandoSeleccionar.AppendLine("		Entregas.Documento, ")
            loComandoSeleccionar.AppendLine("		Entregas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("		Entregas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		Entregas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Renglon, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Precio1, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Por_Des, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas.Mon_Net ")
            loComandoSeleccionar.AppendLine("From Articulos, ")
            loComandoSeleccionar.AppendLine("		Entregas, ")
            loComandoSeleccionar.AppendLine("		Renglones_Entregas, ")
            loComandoSeleccionar.AppendLine("		Clientes, ")
            loComandoSeleccionar.AppendLine("		Vendedores, ")
            loComandoSeleccionar.AppendLine("		Almacenes, ")
            loComandoSeleccionar.AppendLine("		Marcas ")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
            loComandoSeleccionar.AppendLine("		And Renglones_Entregas.Documento = Entregas.Documento ")
            loComandoSeleccionar.AppendLine("		And Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("		And Entregas.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("		And Entregas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("		And Renglones_Entregas.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("		And Articulos.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		And Entregas.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		And Clientes.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		And Vendedores.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		And Marcas.Cod_Mar  between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		And Entregas.Status IN ( " & lcParametro5Desde & " )")
            loComandoSeleccionar.AppendLine("		And Almacenes.Cod_Alm between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("		And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Entregas.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND	Entregas.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro8Hasta)

            'loComandoSeleccionar.AppendLine("ORDER BY Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY  Articulos.Cod_Art, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNentregas_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNentregas_Articulos.ReportSource = loObjetoReporte


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
' MVP: 11/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/04/09: Cambios Estandarización de codigo, correcion campo estatus. 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'