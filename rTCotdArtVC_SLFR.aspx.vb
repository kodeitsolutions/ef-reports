'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCotdArtVC_SLFR"
'-------------------------------------------------------------------------------------------'
Partial Class rTCotdArtVC_SLFR
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

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


            loComandoSeleccionar.AppendLine(" SELECT		Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 				sum (Renglones_Cotizaciones.Can_Art1) as can_art1,")
            loComandoSeleccionar.AppendLine(" 				Renglones_Cotizaciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine(" 				sum (Renglones_Cotizaciones.Mon_Net) as mon_net ")
            loComandoSeleccionar.AppendLine(" FROM			Articulos, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Cotizaciones, ")
            loComandoSeleccionar.AppendLine(" 				Clientes, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores, ")
            loComandoSeleccionar.AppendLine(" 				Monedas, ")
            loComandoSeleccionar.AppendLine(" 				Transportes, ")
            loComandoSeleccionar.AppendLine(" 				Almacenes, ")
            loComandoSeleccionar.AppendLine(" 				Departamentos, ")
            loComandoSeleccionar.AppendLine(" 				Clases_Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE			Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Cotizaciones.Documento = Cotizaciones.Documento ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Mon = Monedas.Cod_Mon")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Cotizaciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <>	'VIA-' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'GEN-' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,5) <> 'NOTAS' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'SES-CAN' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-CAN' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-PER' ")
            'loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'REQ-' ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Cotizaciones.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Departamentos.Cod_Dep  between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Clases_Articulos.Cod_Cla  between" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Almacenes.Cod_Alm between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Mon between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine(" 			    AND Articulos.Cod_Dep IN (" & lcDepartamentosUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			    AND Articulos.Cod_Mar IN (" & lcMarcasUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			    AND Renglones_Cotizaciones.Cod_Alm IN (" & lcAlmacenesUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			    AND Cotizaciones.Cod_Ven IN (" & lcVendedoresUsuario & ")")

            loComandoSeleccionar.AppendLine(" GROUP BY		Cotizaciones.Cod_Ven, Cotizaciones.Cod_Cli, Renglones_Cotizaciones.Cod_Art,  Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("               Renglones_Cotizaciones.Cod_Uni, Vendedores.Nom_Ven, Clientes.Nom_Cli ")
            'loComandoSeleccionar.AppendLine(" ORDER BY		Cotizaciones.Cod_Ven, Cotizaciones.Cod_Cli, Renglones_Cotizaciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("ORDER BY       Cotizaciones.Cod_Ven, Cotizaciones.Cod_Cli, " & lcOrdenamiento)
            'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCotdArtVC_SLFR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCotdArtVC_SLFR.ReportSource = loObjetoReporte


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
' CMS: 12/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JFP: 04/09/12: Adecuacion a nuevo reporte con filtros de la fuerza de venta
'-------------------------------------------------------------------------------------------'

