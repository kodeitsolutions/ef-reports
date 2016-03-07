'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_AnuladasPFechas"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_AnuladasPFechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String 

            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
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
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))

            Dim lcCampoEvaluar As String 

            '------------------------------------------------------'
            ' Ajuste del Parametro Uno Hasta
            '------------------------------------------------------'
            If (lcParametro1Desde = "''") Then
                lcParametro1Hasta = "'zzzzzzz'"
            Else
                lcParametro1Hasta = lcParametro1Desde
            End If

            '------------------------------------------------------'
            ' Ajuste del campo de fecha por el cual evaluar
            '------------------------------------------------------'
            If (lcParametro3Desde = "'Anulado'") Then
                lcCampoEvaluar  =   "'Auditorias.Registro'"
            Else
                lcCampoEvaluar  =   "'Cuentas_Cobrar.Fec_Ini'"
            End If

            lcCampoEvaluar  = lcCampoEvaluar.Replace("'","")



            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Equipo, ")
            loComandoSeleccionar.AppendLine("           Auditorias.Registro, ")
            loComandoSeleccionar.AppendLine("           CAST(DATEPART(HH, Auditorias.Registro) AS VARCHAR) + ':' + CAST(DATEPART(MI, Auditorias.Registro) AS VARCHAR) + ':' + CAST(DATEPART(SS, Auditorias.Registro) AS VARCHAR) AS Hora, " )
            loComandoSeleccionar.AppendLine("           Auditorias.Cod_Usu, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Suc, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Bru * (CASE WHEN Tip_Doc = 'Credito' THEN -1 ELSE 1 END))   AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Imp1 * (CASE WHEN Tip_Doc = 'Credito' THEN -1 ELSE 1 END))  AS  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Net * (CASE WHEN Tip_Doc = 'Credito' THEN -1 ELSE 1 END))   AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Sal * (CASE WHEN Tip_Doc = 'Credito' THEN -1 ELSE 1 END))   AS  Mon_Sal,  ")
            loComandoSeleccionar.AppendLine("  		    SUBSTRING(Cuentas_Cobrar.Comentario,1,150)  AS Comentario ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Auditorias ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      =   Auditorias.Clave2 ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    =   Auditorias.Documento ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status       =   'Anulado' ")
            loComandoSeleccionar.AppendLine("           AND Auditorias.Tabla            =   'Cuentas_Cobrar' ")
            loComandoSeleccionar.AppendLine("           AND Auditorias.Accion           =   'Anular' ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND " & lcCampoEvaluar & "      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon            BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip            BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cla            BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tra      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon      BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 




            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_AnuladasPFechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_AnuladasPFechas.ReportSource = loObjetoReporte

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
' JJD: 07/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
