'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Facturados"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Facturados
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art,					")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Mar,					")
            loComandoSeleccionar.AppendLine(" 			Marcas.Nom_Mar,						")
            loComandoSeleccionar.AppendLine(" 			RTRIM(Facturas.Cod_Ven) + ' ' + SUBSTRING(RTRIM(Vendedores.Nom_Ven),1,30) AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1,		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Uni, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Precio1, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Pro1,		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Ult1,		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Por_Des, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Mon_Des, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Mon_Net			")
            loComandoSeleccionar.AppendLine("FROM		Articulos,							")
            loComandoSeleccionar.AppendLine(" 			Marcas,								")
            loComandoSeleccionar.AppendLine(" 			Facturas,							")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas,					")
            loComandoSeleccionar.AppendLine(" 			Vendedores							")
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Art                   =   Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status					IN  ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Documento    =   Facturas.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven                =   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar  BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla  BETWEEN" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Tip  BETWEEN" & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Rev  BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Ubi  BETWEEN" & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Facturados", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Facturados.ReportSource = loObjetoReporte

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
' YJP:  21/04/09: Estandarizacion de codigo.
'-------------------------------------------------------------------------------------------'
' CMS:  12/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' CMS:  14/07/09: Se agrego filtro Cod_Rev
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro.					' 
'-------------------------------------------------------------------------------------------'
' MAT: 21/12/10: Ajustado el ordenamiento según Código de Artículo y Marca					' 
'-------------------------------------------------------------------------------------------'
