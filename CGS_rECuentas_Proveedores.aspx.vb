﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rECuentas_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rECuentas_Proveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Proveedores.Cod_Pro,     ")
            loComandoSeleccionar.AppendLine("		(SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				  THEN Cuentas_Pagar.Mon_Sal      ")
            loComandoSeleccionar.AppendLine("				  ELSE 0 END)      ")
            loComandoSeleccionar.AppendLine("		- SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				   THEN Cuentas_Pagar.Mon_Sal     ")
            loComandoSeleccionar.AppendLine("				   ELSE 0 END))					AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO #tempSALDOINICIAL           ")
            loComandoSeleccionar.AppendLine("FROM Proveedores     ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro     ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Fec_Ini < @ldFecha_Desde  ")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro    ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 	Cuentas_Pagar.Cod_Pro,     ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("       CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("            THEN Cuentas_Pagar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("            ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Comentario")
            loComandoSeleccionar.AppendLine("INTO #tempMOVIMIENTO    ")
            loComandoSeleccionar.AppendLine("FROM Proveedores    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro    ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro,     ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END							AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("       CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("            ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						    AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Comentario")
            loComandoSeleccionar.AppendLine("FROM Proveedores    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro    ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip <> 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0 ")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("    AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tempMOVIMIENTO.Cod_Pro,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Nom_Pro,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Cod_Tip,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Documento,  ")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Control,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Ini,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Reg,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Factura,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Deb,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Hab,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Comentario, ")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Desde AS DATE) AS Fec_Desde,")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Hasta AS DATE) AS Fec_Hasta,")
            loComandoSeleccionar.AppendLine("		COALESCE(#tempSALDOINICIAL.Sal_Ini,0) AS Sal_Ini,  	")
            loComandoSeleccionar.AppendLine("		0 AS Sal_Doc  	")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO  	")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tempSALDOINICIAL ON #tempSALDOINICIAL.Cod_Pro = #tempMOVIMIENTO.Cod_Pro  ")
            loComandoSeleccionar.AppendLine("ORDER BY #tempMOVIMIENTO.Cod_Pro ASC")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempSALDOINICIAL")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempMOVIMIENTO")


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmente los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Pro", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Pro", GetType(String))
                loColumna.MaxLength = 200
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Documento", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Control", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Reg", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Factura", GetType(String))
                loColumna.MaxLength = 20
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna.MaxLength = 500
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Desde", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Hasta", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Cuenta_Actual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loNuevaFila.Item("Cod_Pro") = loFila("Cod_Pro")
                loNuevaFila.Item("Nom_Pro") = loFila("Nom_Pro")
                loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                loNuevaFila.Item("Documento") = loFila("Documento")
                loNuevaFila.Item("Control") = loFila("Control")
                loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                loNuevaFila.Item("Fec_Reg") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Reg")), "dd/MM/yyyy")
                loNuevaFila.Item("Factura") = loFila("Factura")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Comentario") = loFila("Comentario")
                loNuevaFila.Item("Fec_Desde") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Desde")), "dd/MM/yyyy")
                loNuevaFila.Item("Fec_Hasta") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Hasta")), "dd/MM/yyyy")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Cuenta_Actual = loFila("Cod_Pro")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Pro") <> Cuenta_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Cod_Pro") = loFila("Cod_Pro")
                    loNuevaFila.Item("Nom_Pro") = loFila("Nom_Pro")
                    loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                    loNuevaFila.Item("Documento") = loFila("Documento")
                    loNuevaFila.Item("Control") = loFila("Control")
                    loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                    loNuevaFila.Item("Fec_Reg") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Reg")), "dd/MM/yyyy")
                    loNuevaFila.Item("Factura") = loFila("Factura")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Comentario") = loFila("Comentario")
                    loNuevaFila.Item("Fec_Desde") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Desde")), "dd/MM/yyyy")
                    loNuevaFila.Item("Fec_Hasta") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Hasta")), "dd/MM/yyyy")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Cuenta_Actual = loFila("Cod_Pro")

                    loTabla.AcceptChanges()

                Next lnNumeroFila


                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'
                Me.mCargarLogoEmpresa(loDatosReporteFinal.Tables(0), "LogoEmpresa")
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

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rECuentas_Proveedores", loDatosReporteFinal)
            Else

                ''-------------------------------------------------------------------------------------------------------
                '' Verificando si el select (tabla nº0) trae registros
                ''-------------------------------------------------------------------------------------------------------

                If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                              "No se Encontraron Registros para los Parámetros Especificados. ", _
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                               "350px", _
                                               "200px")
                End If


                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rECuentas_Proveedores", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rECuentas_Proveedores.ReportSource = loObjetoReporte

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
' CMS: 01/07/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS:  31/07/09: Filtro “Revisión:”, verificacion de registro
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'
' MAT:  13/05/11: Ajuste de la Consulta, Tomaba los Saldos de Docuemntos Anulados
'-------------------------------------------------------------------------------------------'