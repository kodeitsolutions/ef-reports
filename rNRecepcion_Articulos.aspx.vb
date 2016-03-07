Imports System.Data
Partial Class rNRecepcion_Articulos
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
            ' Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Status, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Recepciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Almacenes, ")
            loComandoSeleccionar.AppendLine("           Departamentos, ")
            loComandoSeleccionar.AppendLine("           Secciones, ")
            loComandoSeleccionar.AppendLine("           Marcas ")
            loComandoSeleccionar.AppendLine(" WHERE     Recepciones.Documento               =   Renglones_Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("           And Renglones_Recepciones.Cod_Art   =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec               =   Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("           And Secciones.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Recepciones.Cod_Pro             =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Recepciones.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Renglones_Recepciones.Cod_Alm   =   Almacenes.Cod_Alm ")
            'loComandoSeleccionar.AppendLine("           And Renglones_Recepciones.Cod_Art   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And Renglones_Recepciones.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Recepciones.Fec_Ini             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Recepciones.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec               Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar               Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Recepciones.Status              IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Recepciones.Cod_Alm   Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Recepciones.Cod_rev   Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Recepciones.Cod_Suc   Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Renglones_Recepciones.Cod_Art, Recepciones.Fec_Ini, Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    Renglones_Recepciones.Cod_Art,  " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNRecepcion_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNRecepcion_Articulos.ReportSource = loObjetoReporte

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
' CMS: 22/06/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP: 01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Ajuste a la Relacion Departamento-Sucursal
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Se Ajusto el primer filtro de cod_art a documento
'-------------------------------------------------------------------------------------------'