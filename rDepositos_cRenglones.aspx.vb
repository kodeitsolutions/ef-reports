'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDepositos_cRenglones"
'-------------------------------------------------------------------------------------------'
Partial Class rDepositos_cRenglones

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

            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Depositos.Documento, ")
            loComandoSeleccionar.AppendLine("           Depositos.Num_Dep, ")
            loComandoSeleccionar.AppendLine("           Depositos.Status, ")
            loComandoSeleccionar.AppendLine("           Depositos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Mon AS Moneda, ")
            loComandoSeleccionar.AppendLine("           Depositos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Efe, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Che, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Tar, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Otr, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Depositos.Actualiza, ")
            loComandoSeleccionar.AppendLine("           Depositos.Foto, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Depositos.Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Com, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Com, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Ret, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           (Renglones_Depositos.Mon_Com + Renglones_Depositos.Mon_Imp) AS Com_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Tipo, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Referencia, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Net AS  Mon_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING((CASE WHEN Renglones_Depositos.Tipo = 'Efectivo' Then ' ' Else 'Origen: ' + CAST(UPPER(Renglones_Depositos.Tip_Ori) AS VARCHAR(20)) + ' Número: ' + CAST(Renglones_Depositos.Doc_Ori AS VARCHAR(20)) End),1,50) AS Origen ")
            loComandoSeleccionar.AppendLine(" FROM      Depositos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Depositos.Documento     =   Renglones_Depositos.Documento ")
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Cue   =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Con   =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           And Depositos.Documento     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Fec_Ini       Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Num_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Cue       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Con       Between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Mon       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Status        IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_rev       Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Suc       Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'loComandoSeleccionar.AppendLine(" ORDER BY  Renglones_Depositos.Cod_Caj, Renglones_Depositos.Tipo, Depositos.Documento ")

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDepositos_cRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDepositos_cRenglones.ReportSource = loObjetoReporte

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
' CMS: 18/04/09: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' AAP: 01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
