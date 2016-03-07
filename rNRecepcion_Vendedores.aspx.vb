Imports System.Data
Partial Class rNRecepcion_Vendedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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
        Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
        'Dim lcParametro8Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro8Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        'Dim lcParametro9Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro9Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Recepciones.Status, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Nom_Pro,1, 30)                        AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Recepciones.Mon_Imp1 + Recepciones.Mon_Imp2 + Recepciones.Mon_Imp3)    AS Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Sal, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra ")
            loComandoSeleccionar.AppendLine(" FROM      Recepciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Recepciones.Cod_Pro				=   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Ven			=   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Tra			=   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Documento		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Fec_Ini			BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Pro			BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Ven			BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Tra			BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_Mon			BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Recepciones.Status			IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Cod_rev			BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Recepciones.Cod_Ven, Recepciones.Fec_Ini, Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNRecepcion_Vendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNRecepcion_Vendedores.ReportSource = loObjetoReporte

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
' JJD: 14/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Agregar Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'