'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCajas_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rMCajas_Numeros

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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cajas, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cajas.Cod_Con       =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Cod_Caj   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Cod_Mon   Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Cod_rev  Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cajas.Cod_Suc  Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Movimientos_Cajas.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMCajas_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCajas_Numeros.ReportSource = loObjetoReporte

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
' JJD: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 21/02/09: Ajustes al reporte
'-------------------------------------------------------------------------------------------'
' GCR: 26/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'