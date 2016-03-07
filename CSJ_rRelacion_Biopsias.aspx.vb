'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CSJ_rRelacion_Biopsias"
'-------------------------------------------------------------------------------------------'
Partial Class CSJ_rRelacion_Biopsias
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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
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


            loComandoSeleccionar.AppendLine(" SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Status, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Por_Des, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven ")
            loComandoSeleccionar.AppendLine(" FROM			Facturas ")
            loComandoSeleccionar.AppendLine(" JOIN	Renglones_Facturas On (Renglones_Facturas.Documento	= Facturas.Documento)")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Facturas.Cod_Art		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" JOIN	Articulos On (Articulos.Cod_Art	=	Renglones_Facturas.Cod_Art )")
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep       		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Mar				BETWEEN" & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Marcas On (Articulos.Cod_Mar	=	Marcas.Cod_Mar )")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Clientes On (Facturas.Cod_Cli	=	Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Vendedores On (Facturas.Cod_Ven	=	Vendedores.Cod_Ven)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Almacenes On (Renglones_Facturas.Cod_Alm	=	Almacenes.Cod_Alm )")
            loComandoSeleccionar.AppendLine(" 				AND Almacenes.Cod_Alm				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" WHERE			Facturas.Fec_Ini					BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Cli				BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Ven				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Mon				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Tra				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Status					IN (" & lcParametro9Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_For				BETWEEN" & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("               AND Facturas.Cod_Rev between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Suc				BETWEEN" & lcParametro13Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY       Articulos.Cod_Art, " & lcOrdenamiento)


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

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CSJ_rRelacion_Biopsias", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_rRelacion_Biopsias.ReportSource = loObjetoReporte


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
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  04/03/09: Estandarizacion de codigo
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  14/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Se agregaron los siguientes filtros: Departamento, Clase de articulo, 
'                 Transporte, Moneda, Forma de pago, sucursal, verificacion de registro
'-------------------------------------------------------------------------------------------'
' MAT:  22/02/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'
' MAT:  22/02/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'