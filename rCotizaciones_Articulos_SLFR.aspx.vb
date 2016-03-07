'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCotizaciones_Articulos_SLFR"
'-------------------------------------------------------------------------------------------'
Partial Class rCotizaciones_Articulos_SLFR
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
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcUsuario As String = goServicios.mObtenerCampoFormatoSQL(goUsuario.pcCodigo)
            loComandoSeleccionar.AppendLine("SELECT		COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-ALM'       AND Cod_Reg = " & lcUsuario & "), '') AS Almacenes,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-DPTO'      AND Cod_Reg = " & lcUsuario & "), '') AS Departamentos,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-VENDED'    AND Cod_Reg = " & lcUsuario & "), '') AS Vendedores,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-MARCAS'    AND Cod_Reg = " & lcUsuario & "), '') AS Marcas")

            loComandoSeleccionar.AppendLine("")

            Dim loPermisos As DataTable
            loPermisos = (New goDatos()).mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "Campos_Propiedades").Tables(0)
            loComandoSeleccionar.Length = 0

            Dim lcAlmacenesUsuario As String = CStr(loPermisos.Rows(0).Item("Almacenes")).Trim()
            Dim lcDepartamentosUsuario As String = CStr(loPermisos.Rows(0).Item("Departamentos")).Trim()
            Dim lcVendedoresUsuario As String = CStr(loPermisos.Rows(0).Item("Vendedores")).Trim()
            Dim lcMarcasUsuario As String = CStr(loPermisos.Rows(0).Item("Marcas")).Trim()

            lcAlmacenesUsuario = goServicios.mObtenerListaFormatoSQL(lcAlmacenesUsuario)
            lcDepartamentosUsuario = goServicios.mObtenerListaFormatoSQL(lcDepartamentosUsuario)
            lcVendedoresUsuario = goServicios.mObtenerListaFormatoSQL(lcVendedoresUsuario)
            lcMarcasUsuario = goServicios.mObtenerListaFormatoSQL(lcMarcasUsuario)


            loComandoSeleccionar.Append(" SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.Append(" 				Articulos.Nom_art, ")
            loComandoSeleccionar.Append(" 				Articulos.Cod_Mar, ")
            loComandoSeleccionar.Append(" 				Articulos.Status, ")
            loComandoSeleccionar.Append(" 				Cotizaciones.Documento, ")
            loComandoSeleccionar.Append(" 				Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.Append(" 				Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.Append(" 				Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Cod_Alm, ")

            Select Case lcParametro9Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_Cotizaciones.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_Cotizaciones.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_Cotizaciones.Can_Art1 - Renglones_Cotizaciones.Can_Pen1) AS Can_Art1, ")
            End Select

            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Cod_Uni, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Precio1, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Por_Des, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones.Mon_Net ")
            loComandoSeleccionar.Append(" FROM			Articulos, ")
            loComandoSeleccionar.Append(" 				Cotizaciones, ")
            loComandoSeleccionar.Append(" 				Renglones_Cotizaciones, ")
            loComandoSeleccionar.Append(" 				Clientes, ")
            loComandoSeleccionar.Append(" 				Vendedores, ")
            loComandoSeleccionar.Append(" 				Almacenes, ")
            loComandoSeleccionar.Append(" 				Marcas ")
            loComandoSeleccionar.Append(" WHERE			Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art ")
            loComandoSeleccionar.Append(" 				AND Renglones_Cotizaciones.Documento = Cotizaciones.Documento ")
            loComandoSeleccionar.Append(" 				AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.Append(" 				AND Cotizaciones.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.Append(" 				AND Cotizaciones.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.Append(" 				AND Renglones_Cotizaciones.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.Append(" 				AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.Append(" 				AND Cotizaciones.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.Append(" 				AND Cotizaciones.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.Append(" 				AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.Append(" 				AND Marcas.Cod_Mar  BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.Append(" 				AND Cotizaciones.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.Append(" 				AND Almacenes.Cod_Alm BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    		AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    		AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep IN (" & lcDepartamentosUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar IN (" & lcMarcasUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cotizaciones.Cod_Alm IN (" & lcAlmacenesUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Ven IN (" & lcVendedoresUsuario & ")")
            'loComandoSeleccionar.Append(" ORDER BY		Articulos.Cod_Art")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCotizaciones_Articulos_SLFR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCotizaciones_Articulos_SLFR.ReportSource = loObjetoReporte


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
' JJD: 24/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  26/02/09: Estandarizacion de codigo
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  20/05/09: Metodo de ordenamiento, verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 verificacion de registros
'-------------------------------------------------------------------------------------------'
' JFP:  31/08/12: Adecuacion para los filtros de la Fuerza de Ventas
'-------------------------------------------------------------------------------------------'
