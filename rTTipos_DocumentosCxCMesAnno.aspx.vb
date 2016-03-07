'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTTipos_DocumentosCxCMesAnno"
'-------------------------------------------------------------------------------------------'
Partial Class rTTipos_DocumentosCxCMesAnno
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT	Cuentas_Cobrar.Cod_Tip                  AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Documentos.Nom_Tip                AS  Nom_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Tip_Doc                  AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, Cuentas_Cobrar.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("			DATEPART(MONTH, Cuentas_Cobrar.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Sal ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTD ")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine(" 			Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE 	Cuentas_Cobrar.Cod_Tip          =   Tipos_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND ((" & lcParametro6Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("           OR  (" & lcParametro6Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 OR Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status       IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon            BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Rev      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Nom_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Anno, ")
            loComandoSeleccionar.AppendLine(" 			Mes, ")
            loComandoSeleccionar.AppendLine(" 			COUNT(Cod_Tip)   AS  Can_Doc, ")
            loComandoSeleccionar.AppendLine(" 			SUM(CASE WHEN Tip_Doc = 'Credito' THEN (Mon_Bru *(-1)) ELSE Mon_Bru END)    AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			SUM(CASE WHEN Tip_Doc = 'Credito' THEN (Mon_Imp1 *(-1)) ELSE Mon_Imp1 END)  AS  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 			SUM(CASE WHEN Tip_Doc = 'Credito' THEN (Mon_Net *(-1)) ELSE Mon_Net END)    AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Tip_Doc = 'Credito' THEN (Mon_Sal *(-1)) ELSE Mon_Sal END)    AS  Mon_Sal ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpTD ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Tip, Nom_Tip, Anno, Mes ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Tip, Nom_Tip," & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTTipos_DocumentosCxCMesAnno", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTTipos_DocumentosCxCMesAnno.ReportSource = loObjetoReporte

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
' JJD: 16/01/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 16/01/10: Se modifico el orden del ordenamiento.
'-------------------------------------------------------------------------------------------'
' CMS: 30/04/10: Se aplico el metodo de validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'