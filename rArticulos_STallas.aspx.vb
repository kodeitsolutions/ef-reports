'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_STallas"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_STallas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT        Articulos.Cod_Mar, ")
            lcComandoSeleccionar.AppendLine("               Articulos.Modelo, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '34' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla34, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '35' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla35, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '36' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla36, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '37' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla37, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '38' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla38, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '39' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla39, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '40' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla40, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '41' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla41, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '42' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla42, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '43' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla43, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '44' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla44, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '45' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla45, ")
            lcComandoSeleccionar.AppendLine("               SUM(CASE WHEN Articulos.Talla = '46' THEN Articulos.Exi_Act1 ELSE 0 END) AS Talla46 ")
            lcComandoSeleccionar.AppendLine(" FROM          Articulos, ")
            lcComandoSeleccionar.AppendLine("               Marcas, ")
            lcComandoSeleccionar.AppendLine("               Departamentos, ")
            lcComandoSeleccionar.AppendLine("               Secciones, ")
            lcComandoSeleccionar.AppendLine("               Tipos_Articulos, ")
            lcComandoSeleccionar.AppendLine("               Clases_Articulos, ")
            lcComandoSeleccionar.AppendLine("               Proveedores, ")
            lcComandoSeleccionar.AppendLine("               Almacenes ")
            lcComandoSeleccionar.AppendLine(" WHERE         Articulos.Cod_Mar       =   Marcas.Cod_Mar ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Dep   =   Departamentos.Cod_Dep ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Sec   =   Secciones.Cod_Sec ")
            lcComandoSeleccionar.AppendLine(" 				AND Departamentos.Cod_Dep = Secciones.Cod_Dep ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Tip   =   Tipos_Articulos.Cod_Tip ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Cla   =   Clases_Articulos.Cod_Cla ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Pro   =   Proveedores.Cod_Pro ")
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Art   Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Mar   Between " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Dep   Between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Sec   Between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Tip   Between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Cla   Between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Cod_Pro   Between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("               And " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("               And Articulos.Status    IN (" & lcParametro7Desde & ")")
            lcComandoSeleccionar.AppendLine(" GROUP BY      Articulos.Cod_Mar,  Articulos.Modelo ")
            'lcComandoSeleccionar.AppendLine(" ORDER BY      Articulos.Cod_Mar,  Articulos.Modelo ")
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_STallas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_STallas.ReportSource = loObjetoReporte

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
' JJD: 25/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 02/02/09: Ajuste del filtro de Status.
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  10/05/10: Se agrego la restriccion Departamentos.Cod_Dep = Secciones.Cod_Dep
'-------------------------------------------------------------------------------------------'