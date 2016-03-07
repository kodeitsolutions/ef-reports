'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRecibos_Lotes_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rRecibos_Lotes_CCMV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = "'zzzzzzz'"

            If (lcParametro1Desde = "''") Then
                lcParametro1Hasta = "'zzzzzzz'"
            Else
                lcParametro1Hasta = lcParametro1Desde 
            End If

            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = "'zzzzzzz'"

            If (lcParametro4Desde = "''") Then
                lcParametro4Hasta = "'zzzzzzz'"
            Else
                lcParametro4Hasta = lcParametro4Desde 
            End If

            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Comentario, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")

            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)   AS  Mon_Gra, ")

            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           CAST('' AS CHAR(400))                               AS  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           DBO.CONCATENA(Cuentas_Cobrar.Cod_Cli) As Asociados ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli " )
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Fec_Ini      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Status       IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Suc      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)


            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Construccion del monto en letras                 '
            '--------------------------------------------------'
            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros                                                   '
            '-------------------------------------------------------------------------------------------------------'
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                          vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRecibos_Lotes_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRecibos_Lotes_CCMV.ReportSource = loObjetoReporte

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
' JJD: 06/05/10: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
