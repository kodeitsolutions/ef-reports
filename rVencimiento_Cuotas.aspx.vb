'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVencimiento_Cuotas"
'-------------------------------------------------------------------------------------------'
Partial Class rVencimiento_Cuotas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Cobrar.Documento                                                                            AS  Documento, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Tip                                                                              AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Ini                                                                              AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Fin                                                                              AS  Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Cli                                                                              AS  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Fin, 103)					                                    AS	Fecha02, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli                                                                                    AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Dir_Fis                                                                                    AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Ven                                                                              AS  Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Tra                                                                              AS  Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Mon                                                                              AS  Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Control                                                                              AS  Control, ")
            loComandoSeleccionar.AppendLine("           (Case When Tip_Doc = 'Credito' Then Cuentas_Cobrar.Mon_Bru *(-1) Else Cuentas_Cobrar.Mon_Bru End)   AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Case When Tip_Doc = 'Credito' Then Cuentas_Cobrar.Mon_Imp1 *(-1) Else Cuentas_Cobrar.Mon_Imp1 End) AS  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (Case When Tip_Doc = 'Credito' Then Cuentas_Cobrar.Mon_Net *(-1) Else Cuentas_Cobrar.Mon_Net End)   AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Case When Tip_Doc = 'Credito' Then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End)   AS  Mon_Sal, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Art                                                                                   AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art                                                                                   AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Informacion                                                                               AS  Informacion ")
            loComandoSeleccionar.AppendLine(" FROM		Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Transportes, ")
            loComandoSeleccionar.AppendLine(" 			Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE		Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           ANd Cuentas_Cobrar.Cod_Cli      =   Articulos.Item ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Mon      =   Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Mon_Sal      >   0 ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Fin      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status       IN  ( " & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")



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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVencimiento_Cuotas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrVencimiento_Cuotas.ReportSource = loObjetoReporte

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
' JJD: 26/03/10: Programacion inicial
'-------------------------------------------------------------------------------------------'