'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNRecepcion_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rNRecepcion_Fechas
    Inherits vis2Formularios.frmReporte

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
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
		Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
        'Dim lcParametro9Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro9Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append(" SELECT    Recepciones.Status, ")
            loComandoSeleccionar.Append("           Recepciones.Documento, ")
            loComandoSeleccionar.Append("           Recepciones.Cod_Pro, ")
            loComandoSeleccionar.Append("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.Append("           Recepciones.Fec_Ini, ")
            loComandoSeleccionar.Append("           Recepciones.Cod_Ven, ")
            loComandoSeleccionar.Append("           Recepciones.Mon_Bru, ")
            loComandoSeleccionar.Append("           (Recepciones.Mon_Imp1 + Recepciones.Mon_Imp2 + Recepciones.Mon_Imp3) AS Mon_Imp, ")
            loComandoSeleccionar.Append("           Recepciones.Mon_Net, ")
            loComandoSeleccionar.Append("           Recepciones.Mon_Sal, ")
            loComandoSeleccionar.Append("           Recepciones.Cod_Tra, ")
            loComandoSeleccionar.Append("           Transportes.Nom_Tra ")
            loComandoSeleccionar.Append(" FROM      Recepciones, ")
            loComandoSeleccionar.Append("           Proveedores, ")
            loComandoSeleccionar.Append("           Transportes, ")
            loComandoSeleccionar.Append("           Vendedores ")
            loComandoSeleccionar.Append(" WHERE     Recepciones.Cod_Pro                 =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Tra             =   Transportes.Cod_Tra ")
            loComandoSeleccionar.Append("           AND Recepciones.Documento           BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Fec_Ini             BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Pro             BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Ven             BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Status              In ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Tra             BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Mon             BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Cod_rev             BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.Append("           AND Recepciones.Cod_Suc             BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.Append("           AND " & lcParametro8Hasta)

            'loComandoSeleccionar.Append(" ORDER BY  Recepciones.Fec_Ini, Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    CONVERT(nchar(30), Recepciones.Fec_Ini,112) DESC," & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNRecepcion_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrNRecepcion_Fechas.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
