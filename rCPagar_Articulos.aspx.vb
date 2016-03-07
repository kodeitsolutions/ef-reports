Imports System.Data
Partial Class rCPagar_Articulos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
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

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Pagar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Status, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Renglon                AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Art                AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Precio1                AS  Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Can_Art                AS  Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Uni                AS  Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Mon_Net                AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Comentario             AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,50)           AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,30)          AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,30)         AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)        AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Departamentos.Nom_Dep,1,30)       AS  Nom_Dep, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Secciones.Nom_Sec,1,30)           AS  Nom_Sec, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Marcas.Nom_Mar,1,30)              AS  Nom_Mar, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Nom_Pro,1,50)         AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Tip                AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Tipos_Documentos.Nom_Tip,1,40)    AS  Nom_Tip ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Departamentos, ")
            loComandoSeleccionar.AppendLine("           Secciones, ")
            loComandoSeleccionar.AppendLine("           Marcas, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Documento             =   Renglones_Documentos.Documento ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip           =   Renglones_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip           =   Tipos_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro           =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven           =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_For           =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tra           =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art               =   Renglones_Documentos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Secciones.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec               =   Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("           And Renglones_Documentos.Origen     =   'Compras' ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Documento        Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Fec_Ini          Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven          Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Status           IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Documentos.Cod_Art    Between" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep               Between" & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec               Between" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar               Between" & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
             loComandoSeleccionar.AppendLine("          And Cuentas_Pagar.Cod_rev              Between" & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Renglones_Documentos.Cod_Art, Renglones_Documentos.Cod_Tip, Renglones_Documentos.Documento ")
             loComandoSeleccionar.AppendLine("ORDER BY   Renglones_Documentos.Cod_Art, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_Articulos.ReportSource = loObjetoReporte

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
' JJD: 07/03/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'