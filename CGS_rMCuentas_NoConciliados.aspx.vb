'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rMCuentas_NoConciliados"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rMCuentas_NoConciliados

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
           
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN  Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine(" JOIN Ordenes_Pagos ON Movimientos_Cuentas.Doc_Ori = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Tip_Ori = 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Doc_Ori = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Doc_Cil     =   '0' ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN  Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine(" JOIN Pagos ON Movimientos_Cuentas.Doc_Ori = Pagos.Documento")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Doc_Ori = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Doc_Cil     =   '0' ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Movimientos_Cuentas.Cod_Tip = 'Cre-TP' ")
            loComandoSeleccionar.AppendLine("           THEN 'Transferencia' ELSE Movimientos_Cuentas.Tipo END AS Tipo,")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Comentario AS Nom_Pro")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN  Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Doc_Cil     =   '0' ")
            loComandoSeleccionar.AppendLine("           AND (Movimientos_Cuentas.Cod_Tip = 'Cre-TP' OR Movimientos_Cuentas.Cod_Tip = 'Cre-NC')")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY   Movimientos_Cuentas.Cod_Cue, " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rMCuentas_NoConciliados", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCuentas_NoConciliados.ReportSource = loObjetoReporte

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
' JJD: 21/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT:  06/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT:  29/09/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'