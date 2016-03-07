'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTFacArtVC"
'-------------------------------------------------------------------------------------------'
Partial Class rTFacArtVC
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 				sum (Renglones_Facturas.Can_Art1) as can_art1,")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine(" 				sum (Renglones_Facturas.Mon_Net) as mon_net ")
            loComandoSeleccionar.AppendLine(" FROM			Articulos, ")
            loComandoSeleccionar.AppendLine(" 				Facturas, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine(" 				Clientes, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores, ")
            loComandoSeleccionar.AppendLine(" 				Monedas, ")
            loComandoSeleccionar.AppendLine(" 				Transportes, ")
            loComandoSeleccionar.AppendLine(" 				Almacenes, ")
            loComandoSeleccionar.AppendLine(" 				Departamentos, ")
            loComandoSeleccionar.AppendLine(" 				Clases_Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE			Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Facturas.Documento = Facturas.Documento ")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Mon = Monedas.Cod_Mon")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Facturas.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <>	'VIA-' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'GEN-' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,5) <> 'NOTAS' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'SES-CAN' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-CAN' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-PER' ")
            loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'REQ-' ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Facturas.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Departamentos.Cod_Dep  between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Clases_Articulos.Cod_Cla  between" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Almacenes.Cod_Alm between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Mon between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("               AND Facturas.Cod_Rev between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("               AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY		Facturas.Cod_Ven, Facturas.Cod_Cli, Renglones_Facturas.Cod_Art,  Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Cod_Uni, Vendedores.Nom_Ven, Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("ORDER BY      Facturas.Cod_Ven , Facturas.Cod_Cli, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Facturas.Cod_Ven , Facturas.Cod_Cli, Renglones_Facturas.Cod_Art ")
            

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTFacArtVC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTFacArtVC.ReportSource = loObjetoReporte


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
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  02/07/09 : Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  18/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'