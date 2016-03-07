'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rECuentas_Cajas_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class rECuentas_Cajas_Stratos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Cod_Caj,     ")
            loComandoSeleccionar.AppendLine("			(SUM(Movimientos_Cajas.Mon_Deb)- SUM(Movimientos_Cajas.Mon_Hab)) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO		#tempSALDOINICIAL     ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas     ")
             loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cajas.Fec_Ini < " & lcParametro1Desde & "  ")
            loComandoSeleccionar.AppendLine("	        AND		Movimientos_Cajas.Cod_Caj BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("	        AND		" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Ban       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Con       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Status        IN ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev between " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev NOT between " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tipo        IN ( " & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cajas.Cod_Caj     ")


            loComandoSeleccionar.AppendLine(" SELECT	Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Comentario,")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine(" INTO		#tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine(" JOIN      Cajas ON  Movimientos_Cajas.Cod_Caj         =   Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine(" JOIN      Conceptos ON Movimientos_Cajas.Cod_Con      =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cajas.Cod_Caj       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Ban       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Con       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Status        IN ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev BETWEEN " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev NOT BETWEEN " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tipo        IN ( " & lcParametro8Desde & ")")


            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Documento,   	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Tip_Doc,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Cod_Caj,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Nom_Caj,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Cod_Ban,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Fec_Ini,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Status,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Referencia,    	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Cod_Con,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Nom_Con,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Cod_Mon,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Mon_Deb,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Mon_Hab,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Tip_Ori,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Doc_Ori,     	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Comentario,  	")
            loComandoSeleccionar.AppendLine("    	#tempMOVIMIENTO.Tipo,		    ")
            loComandoSeleccionar.AppendLine("    	ISNULL(#tempSALDOINICIAL.Sal_Ini,0) AS Sal_Ini,	")
            loComandoSeleccionar.AppendLine("    	0 AS Sal_Doc    ")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("LEFT JOIN	#tempSALDOINICIAL ON #tempSALDOINICIAL.Cod_Caj = #tempMOVIMIENTO.Cod_Caj")
            loComandoSeleccionar.AppendLine("ORDER BY     #tempMOVIMIENTO." & lcOrdenamiento & ", #tempMOVIMIENTO.Fec_Ini ASC")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmetne los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Documento", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Tip_Doc", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Caj", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Caj", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Ban", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Status", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Referencia", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Con", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Con", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Mon", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Tip_Ori", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Doc_Ori", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Tipo", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna.MaxLength = 500
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Caja_Actual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loColumna = New DataColumn("Documento", GetType(String))
                loColumna = New DataColumn("Tip_Doc", GetType(String))
                loColumna = New DataColumn("Cod_Caj", GetType(String))
                loColumna = New DataColumn("Nom_Caj", GetType(String))
                loColumna = New DataColumn("Cod_Ban", GetType(String))
                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loColumna = New DataColumn("Status", GetType(String))
                loColumna = New DataColumn("Referencia", GetType(String))
                loColumna = New DataColumn("Cod_Con", GetType(String))
                loColumna = New DataColumn("Nom_Con", GetType(String))
                loColumna = New DataColumn("Cod_Mon", GetType(String))
                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loColumna = New DataColumn("Tip_Ori", GetType(String))
                loColumna = New DataColumn("Doc_Ori", GetType(String))
                loColumna = New DataColumn("Tipo", GetType(String))
                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))

                loNuevaFila.Item("Documento") = loFila("Documento")
                loNuevaFila.Item("Tip_Doc") = loFila("Tip_Doc")
                loNuevaFila.Item("Cod_Caj") = loFila("Cod_Caj")
                loNuevaFila.Item("Nom_Caj") = loFila("Nom_Caj")
                loNuevaFila.Item("Cod_Ban") = loFila("Cod_Ban")
                loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "MM/dd/yyyy")
                loNuevaFila.Item("Status") = loFila("Status")
                loNuevaFila.Item("Referencia") = loFila("Referencia")
                loNuevaFila.Item("Cod_Con") = loFila("Cod_Con")
                loNuevaFila.Item("Nom_Con") = loFila("Nom_Con")
                loNuevaFila.Item("Cod_Mon") = loFila("Cod_Mon")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Comentario") = loFila("Comentario")
                loNuevaFila.Item("Tip_Ori") = loFila("Tip_Ori")
                loNuevaFila.Item("Doc_Ori") = loFila("Doc_Ori")
                loNuevaFila.Item("Tipo") = loFila("Tipo")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

              
                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Caja_Actual = loFila("Cod_Caj")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Caj") <> Caja_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If


                    loNuevaFila.Item("Documento") = loFila("Documento")
                    loNuevaFila.Item("Tip_Doc") = loFila("Tip_Doc")
                    loNuevaFila.Item("Cod_Caj") = loFila("Cod_Caj")
                    loNuevaFila.Item("Nom_Caj") = loFila("Nom_Caj")
                    loNuevaFila.Item("Cod_Ban") = loFila("Cod_Ban")
                    loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "MM/dd/yyyy")
                    loNuevaFila.Item("Status") = loFila("Status")
                    loNuevaFila.Item("Referencia") = loFila("Referencia")
                    loNuevaFila.Item("Cod_Con") = loFila("Cod_Con")
                    loNuevaFila.Item("Nom_Con") = loFila("Nom_Con")
                    loNuevaFila.Item("Cod_Mon") = loFila("Cod_Mon")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Comentario") = loFila("Comentario")
                    loNuevaFila.Item("Tip_Ori") = loFila("Tip_Ori")
                    loNuevaFila.Item("Doc_Ori") = loFila("Doc_Ori")
                    loNuevaFila.Item("Tipo") = loFila("Tipo")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                   
                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Caja_Actual = loFila("Cod_Caj")

                    loTabla.AcceptChanges()

                Next lnNumeroFila

                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '---------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '---------------------------------------------------------------------------------------'

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rECuentas_Cajas_Stratos", loDatosReporteFinal)
            Else
                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rECuentas_Cajas_Stratos", laDatosReporte)
            End If


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrECuentas_Cajas_Stratos.ReportSource = loObjetoReporte

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
' MAT:  08/09/11 : Código Inicial
'-------------------------------------------------------------------------------------------'
