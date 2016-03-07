'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLCompras_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rLCompras_Articulos

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.Append("           Articulos.Nom_Art, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Sec, ")
            loComandoSeleccionar.Append("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.Append("           Libres_Compras.Status, ")
            loComandoSeleccionar.Append("           Libres_Compras.Documento, ")
            loComandoSeleccionar.Append("           Libres_Compras.Cod_Pro, ")
            loComandoSeleccionar.Append("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.Append("           Libres_Compras.Fec_Ini, ")
            loComandoSeleccionar.Append("           Libres_Compras.Cod_Ven, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Renglon, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Cod_Alm, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Can_Art1, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Cod_Uni, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Precio1, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Por_Des, ")
            loComandoSeleccionar.Append("           Renglones_LCompras.Mon_Net ")
            loComandoSeleccionar.Append(" FROM      Articulos, ")
            loComandoSeleccionar.Append("           Libres_Compras, ")
            loComandoSeleccionar.Append("           Renglones_LCompras, ")
            loComandoSeleccionar.Append("           Proveedores, ")
            loComandoSeleccionar.Append("           Vendedores, ")
            loComandoSeleccionar.Append("           Almacenes, ")
            loComandoSeleccionar.Append("           Departamentos, ")
            loComandoSeleccionar.Append("           Secciones, ")
            loComandoSeleccionar.Append("           Marcas ")
            loComandoSeleccionar.Append(" WHERE     Libres_Compras.Documento            =   Renglones_LCompras.Documento ")
            loComandoSeleccionar.Append("           And Renglones_LCompras.Cod_Art      =   Articulos.Cod_Art ")
            loComandoSeleccionar.Append("           And Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.Append("           And Articulos.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.Append("           And Articulos.Cod_Sec               =   Secciones.Cod_Sec ")
            loComandoSeleccionar.Append("           And Secciones.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.Append("           And Libres_Compras.Cod_Pro          =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.Append("           And Libres_Compras.Cod_Ven          =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.Append("           And Renglones_LCompras.Cod_Alm      =   Almacenes.Cod_Alm ")
            loComandoSeleccionar.Append("           And Renglones_LCompras.Cod_Art      Between " & lcParametro0Desde)
            loComandoSeleccionar.Append("           And " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           And Libres_Compras.Fec_Ini          Between " & lcParametro1Desde)
            loComandoSeleccionar.Append("           And " & lcParametro1Hasta)
            loComandoSeleccionar.Append("           And Libres_Compras.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.Append("           And " & lcParametro2Hasta)
            loComandoSeleccionar.Append("           And Libres_Compras.Cod_Ven          Between " & lcParametro3Desde)
            loComandoSeleccionar.Append("           And " & lcParametro3Hasta)
            loComandoSeleccionar.Append("           And Articulos.Cod_Dep               Between " & lcParametro4Desde)
            loComandoSeleccionar.Append("           And " & lcParametro4Hasta)
            loComandoSeleccionar.Append("           And Articulos.Cod_Sec               Between " & lcParametro5Desde)
            loComandoSeleccionar.Append("           And " & lcParametro5Hasta)
            loComandoSeleccionar.Append("           And Articulos.Cod_Mar               Between " & lcParametro6Desde)
            loComandoSeleccionar.Append("           And " & lcParametro6Hasta)
            loComandoSeleccionar.Append("           And Libres_Compras.Status           IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.Append("           And Renglones_LCompras.Cod_Alm      Between " & lcParametro8Desde)
            loComandoSeleccionar.Append("           And " & lcParametro8Hasta)
            loComandoSeleccionar.Append("			And Libres_Compras.Cod_rev      Between " & lcParametro9Desde)
            loComandoSeleccionar.Append("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'loComandoSeleccionar.Append(" ORDER BY  Renglones_LCompras.Cod_Art, Libres_Compras.Fec_Ini, Libres_Compras.Documento ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLCompras_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLCompras_Articulos.ReportSource = loObjetoReporte

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
' JJD: 08/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Ajuste a la relacion Departamento - Seccion
'-------------------------------------------------------------------------------------------'
