'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rGuias_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rGuias_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        'Dim lcParametro8Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro8Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        'Dim lcParametro9Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro9Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Guias.Documento, ")
            loComandoSeleccionar.AppendLine("           Guias.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Guias.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Guias.Control, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine("           Guias, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE     Guias.Cod_Cli       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Ven   =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Tra   =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Mon   =   Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine("           AND Guias.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Fec_Ini   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Cli   BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Ven   BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Status    IN ( " & lcParametro4Desde & " ) ")
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Tra   BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Mon   BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Guias.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Guias.Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("           Guias.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY     CONVERT(nchar(30), Guias.Fec_Ini,112), " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGuias_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrGuias_Fechas.ReportSource = loObjetoReporte


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
' JJD: 20/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  26/06/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Se ajusto el metodo de ordenamiento y se aplico lafuncion convert a la 
'                   fecha en la clausula order by
'-------------------------------------------------------------------------------------------'