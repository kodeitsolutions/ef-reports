'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Articulos 
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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = cusAplicacion.goReportes.paParametrosIniciales(13)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT      Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("             Articulos.Status, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("             SUBSTRING(Clientes.Nom_Cli,1,16) AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Cod_Alm, ")

            Select Case lcParametro13Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_Pedidos.Can_Art1 - Renglones_Pedidos.Can_Pen1) AS Can_Art1, ")
            End Select

            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Por_Des, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Mon_Net ")
            loComandoSeleccionar.AppendLine(" From        Articulos, ")
            loComandoSeleccionar.AppendLine("             Pedidos, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("             Clientes, ")
            loComandoSeleccionar.AppendLine("             Vendedores, ")
            loComandoSeleccionar.AppendLine("             Almacenes, ")
            loComandoSeleccionar.AppendLine("             Marcas ")
            loComandoSeleccionar.AppendLine(" WHERE       Articulos.Cod_Art = Renglones_Pedidos.Cod_Art ")
            loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Documento = Pedidos.Documento ")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Cod_Alm = Almacenes.Cod_Alm ")

            Select Case lcParametro13Desde
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Can_Pen1 <> 0 ")
            End Select

            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             AND Clientes.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("             AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("             AND articulos.Cod_Mar  BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Status IN ( " & lcParametro5Desde & " )")
            loComandoSeleccionar.AppendLine("             AND Clientes.Cod_Cla BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("             AND Clientes.Cod_Zon BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Dep BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Sec BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("             AND Almacenes.Cod_Alm BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Rev between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("    	      AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Suc between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    	      AND " & lcParametro12Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY   Articulos.Cod_Art, " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Articulos.ReportSource = loObjetoReporte


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
' JJD: 19/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  16/04/09: Cambios Estandarización de codigo, correcion campo estatus. 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 Metodo de ordenamiento, Verificacion  de registros
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Se Agrego la restricción Renglones_Pedidos.Can_Pen1 <> 0 cuando el filtro 
'                   BackOrder = BackOrder
'-------------------------------------------------------------------------------------------'