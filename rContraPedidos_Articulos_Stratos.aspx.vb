'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rContraPedidos_Articulos_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class rContraPedidos_Articulos_Stratos
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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT      Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("             Articulos.Status, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("             Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Can_Pen1 AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Por_Des, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM        Articulos, ")
            loComandoSeleccionar.AppendLine("             Pedidos, ")
            loComandoSeleccionar.AppendLine("             Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("             Clientes, ")
            loComandoSeleccionar.AppendLine("             Vendedores, ")
            loComandoSeleccionar.AppendLine("             Almacenes, ")
            loComandoSeleccionar.AppendLine("             Departamentos, ")
            loComandoSeleccionar.AppendLine("             Secciones, ")
            loComandoSeleccionar.AppendLine("             Marcas ")
            loComandoSeleccionar.AppendLine(" WHERE       Articulos.Cod_Art                         =   Renglones_Pedidos.Cod_Art ")
            loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Documento           =   Pedidos.Documento ")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Mar                     =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Dep                     =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Sec                     =   Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("             AND Secciones.Cod_Dep                     =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Cli                       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("             AND Pedidos.Cod_Ven                       =   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Cod_Alm             =   Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Can_Pen1            >=  1 ")
            loComandoSeleccionar.AppendLine("             AND SUBSTRING(Pedidos.Comentario,1,10)    =   'BACK ORDER' ")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Art                     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Fec_Ini                       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             AND Clientes.Cod_Cli                      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("             AND Vendedores.Cod_Ven                    BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Mar                     BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("             AND Pedidos.Status                        IN ( " & lcParametro5Desde & " )")
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Cla                     BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("             AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("             AND Clientes.Cod_Zon                      BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	      AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Dep                     BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	      AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("             AND Articulos.Cod_Sec                     BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	      AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY   Articulos.Cod_Art, " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rContraPedidos_Articulos_Stratos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrContraPedidos_Articulos_Stratos.ReportSource = loObjetoReporte


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
' JJD: 29/03/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
