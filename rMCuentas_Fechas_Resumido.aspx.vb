'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCuentas_Fechas_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rMCuentas_Fechas_Resumido

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
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            
            Dim lcParametro8Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(8)).Trim()
            Dim lcOrigen		  As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(9)).Trim()
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(lcOrigen)

            lcParametro4Hasta = IIf(lcParametro4Hasta = "'zzzzzzz'", "'zzzzzzz'", lcParametro4Desde)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Tipos_Movimientos.Nom_Tip,")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cuentas, ")
            loComandoSeleccionar.AppendLine("           Tipos_Movimientos, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Cod_Tip         =   Tipos_Movimientos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Con     =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Documento   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Mon     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Referencia  BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Status      IN  (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cuentas.cod_rev  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cuentas.Cod_Suc  BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
            If (lcParametro8Desde = "Si") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Automatico = 0")
            End If
            If (lcOrigen <> "") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Tip_Ori = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en curReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMCuentas_Fechas_Resumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCuentas_Fechas_Resumido.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 21/02/09: Codigo inicial (A partir de rMCuentas_Fechas).								'
'-------------------------------------------------------------------------------------------'
' RJG: 28/02/09: Ajuste de layout. Ajuste de ordenamiento de fechas por parámetros.			'
'-------------------------------------------------------------------------------------------'
