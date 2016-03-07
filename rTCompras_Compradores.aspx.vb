'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCompras_Compradores"
'-------------------------------------------------------------------------------------------'
Partial Class rTCompras_Compradores
    Inherits vis2Formularios.frmReporte

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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- NOTA: Al Optimizar el reporte agrequé el nombre 'Temporal' para")
            loConsulta.AppendLine("-- no cambiar el ORDER BY de la definición del reporte")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Temporal.*")
            loConsulta.AppendLine("FROM (")
            loConsulta.AppendLine("SELECT  Vendedores.Cod_Ven                  AS Cod_Ven,")
            loConsulta.AppendLine("        Vendedores.Nom_Ven                  AS Nom_Ven,")
            loConsulta.AppendLine("        COUNT(DISTINCT Compras.Documento)   AS Documentos,")
            loConsulta.AppendLine("        SUM(Renglones_Compras.Can_Art1)     AS Can_Art,")
            loConsulta.AppendLine("        SUM(Renglones_Compras.Mon_Net)      AS Mon_Net,")
            loConsulta.AppendLine("        CAST(" & goOpciones.pnDecimalesParaCantidad() & " AS INT) AS NumDecCant")
            loConsulta.AppendLine("FROM    Compras")
            loConsulta.AppendLine("    JOIN Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            loConsulta.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Compras.Cod_Ven")
            loConsulta.AppendLine("WHERE   Compras.Status <> 'Anulado'")
            loConsulta.AppendLine("    AND Compras.Fec_Ini BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loConsulta.AppendLine("    AND Compras.Cod_Pro BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loConsulta.AppendLine("    AND Compras.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loConsulta.AppendLine("    AND Renglones_Compras.Cod_Art BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loConsulta.AppendLine("    AND Compras.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loConsulta.AppendLine("    AND Compras.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loConsulta.AppendLine("GROUP BY Vendedores.Cod_Ven, Vendedores.Nom_Ven")
            loConsulta.AppendLine(") Temporal")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCompras_Compradores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCompras_Compradores.ReportSource = loObjetoReporte

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
' RJG: 05/12/14: Código Inicial, a partir de rTCompras_Proveedores.                         '
'-------------------------------------------------------------------------------------------'
