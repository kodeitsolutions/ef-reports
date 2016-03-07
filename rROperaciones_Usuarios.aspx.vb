Imports System.Data
Imports cusAplicacion

Partial Class rROperaciones_Usuarios
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar1 As New StringBuilder()
            Dim loComandoSeleccionar2 As New StringBuilder()

            loComandoSeleccionar1.AppendLine("SELECT		RTRIM(Cod_Usu) AS Cod_usu,")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'cobros' THEN 1 ELSE 0 END)) As cobros,")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'entregas' THEN 1 ELSE 0 END)) As not_ent,")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'guias' THEN 1 ELSE 0 END)) As not_des,   ")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'depositos' THEN 1 ELSE 0 END)) As depositos,")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'cuentas_cobrar' THEN 1 ELSE 0 END)) As c_cobrar,")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'movimientos_bancos' THEN 1 ELSE 0 END)) As mov_ban,	")
            loComandoSeleccionar1.AppendLine("				SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'ordenes_pagos' THEN 1 ELSE 0 END)) As ord_pago,	")
            loComandoSeleccionar1.AppendLine("				  SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'cobros' THEN 1 ELSE 0 END)) ")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'entregas' THEN 1 ELSE 0 END))	")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'guias' THEN 1 ELSE 0 END))  ")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'depositos' THEN 1 ELSE 0 END)) ")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'cuentas_cobrar' THEN 1 ELSE 0 END))  ")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'movimientos_bancos' THEN 1 ELSE 0 END)) ")
            loComandoSeleccionar1.AppendLine("				+ SUM((CASE WHEN ADM_AUDITORIAS.tabla = 'ordenes_pagos' THEN 1 ELSE 0 END)) ")
            loComandoSeleccionar1.AppendLine("				AS total  ")
            loComandoSeleccionar1.AppendLine("FROM			Auditorias AS ADM_AUDITORIAS ")
            loComandoSeleccionar1.AppendLine("WHERE ")
            loComandoSeleccionar1.AppendLine("               ADM_AUDITORIAS.registro   Between " & lcParametro0Desde)
            loComandoSeleccionar1.AppendLine("               AND " & lcParametro0Hasta)
            loComandoSeleccionar1.AppendLine("               AND ADM_AUDITORIAS.Cod_Emp Between " & lcParametro2Desde)
            loComandoSeleccionar1.AppendLine("               AND " & lcParametro2Hasta)
            loComandoSeleccionar1.AppendLine("GROUP BY Cod_Usu")
            loComandoSeleccionar1.AppendLine("ORDER BY Cod_Usu")

            loComandoSeleccionar2.AppendLine("SELECT        RTRIM(GLOBAL_USUARIOS.Cod_Usu) AS Cod_usu,  ")
            loComandoSeleccionar2.AppendLine("	  	        GLOBAL_USUARIOS.nom_usu")
            loComandoSeleccionar2.AppendLine("FROM			Usuarios AS GLOBAL_USUARIOS ")
            loComandoSeleccionar2.AppendLine("WHERE ")
            loComandoSeleccionar2.AppendLine("               GLOBAL_USUARIOS.nom_usu  Between " & lcParametro1Desde)
            loComandoSeleccionar2.AppendLine("               AND " & lcParametro1Hasta)
            loComandoSeleccionar2.AppendLine("GROUP BY GLOBAL_USUARIOS.cod_usu, GLOBAL_USUARIOS.nom_usu   ")
            loComandoSeleccionar2.AppendLine("ORDER BY " & lcOrdenamiento)

            Dim loServiciosAuditoria As New cusDatos.goDatos
            Dim laDatosReporteAuditoria As DataSet = loServiciosAuditoria.mObtenerTodosSinEsquema(loComandoSeleccionar1.ToString, "curReportes")

            Dim loServiciosUsuarios As New cusDatos.goDatos
            goDatos.pcNombreAplicativoExterno = "Framework"
            Dim laDatosReporteUsuarios As DataSet = loServiciosUsuarios.mObtenerTodosSinEsquema(loComandoSeleccionar2.ToString, "curReportes")

            Dim loTabla As New DataTable("curReportes")
            Dim loColumna As DataColumn

            loColumna = New DataColumn("Cod_Usu", GetType(String))
            loColumna.MaxLength = 50
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Nom_Usu", GetType(String))
            loColumna.MaxLength = 200
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Cobros", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Not_Ent", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Not_Des", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Depositos", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("C_Cobrar", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Mov_Ban", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Ord_Pago", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            loColumna = New DataColumn("Total", GetType(Decimal))
            loTabla.Columns.Add(loColumna)

            Dim loNuevaFila As DataRow
            Dim lnTotalFilasAuditorias As Integer = laDatosReporteAuditoria.Tables(0).Rows.Count
            Dim lnTotalFilasUsuarios As Integer = laDatosReporteUsuarios.Tables(0).Rows.Count
            Dim loFilaAuditoria As DataRow
            Dim loFilaUsuarios As DataRow

            For lnNumeroFilaAuditoria As Integer = 1 To lnTotalFilasAuditorias - 1

                loFilaAuditoria = laDatosReporteAuditoria.Tables(0).Rows(lnNumeroFilaAuditoria)

                For lnNumeroFilaUsuarios As Integer = 1 To lnTotalFilasUsuarios - 1

                    loFilaUsuarios = laDatosReporteUsuarios.Tables(0).Rows(lnNumeroFilaUsuarios)

                    If loFilaAuditoria("Cod_Usu") = loFilaUsuarios("Cod_Usu") Then

                        loNuevaFila = loTabla.NewRow()
                        loTabla.Rows.Add(loNuevaFila)

                        loNuevaFila.Item("Cod_Usu") = loFilaAuditoria("Cod_Usu")
                        loNuevaFila.Item("Nom_Usu") = loFilaUsuarios("Nom_Usu")
                        loNuevaFila.Item("Cobros") = loFilaAuditoria("Cobros")
                        loNuevaFila.Item("Not_Ent") = loFilaAuditoria("Cobros")
                        loNuevaFila.Item("Not_Des") = loFilaAuditoria("Not_Des")
                        loNuevaFila.Item("Depositos") = loFilaAuditoria("Depositos")
                        loNuevaFila.Item("C_Cobrar") = loFilaAuditoria("C_Cobrar")
                        loNuevaFila.Item("Mov_Ban") = loFilaAuditoria("Mov_Ban")
                        loNuevaFila.Item("Ord_Pago") = loFilaAuditoria("Ord_Pago")
                        loNuevaFila.Item("Total") = loFilaAuditoria("Total")
                        loTabla.AcceptChanges()
                        Exit For
                    End If


                Next lnNumeroFilaUsuarios
            Next lnNumeroFilaAuditoria

            Dim loDatosReporteFinal As New DataSet("curReportes")
            loDatosReporteFinal.Tables.Add(loTabla)

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            If (loDatosReporteFinal.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rROperaciones_Usuarios", loDatosReporteFinal)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrROperaciones_Usuarios.ReportSource = loObjetoReporte

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
' YJP: 25/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 03/08/09: Se reescribio la consulta
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
