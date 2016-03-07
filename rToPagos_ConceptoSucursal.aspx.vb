'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rToPagos_ConceptoSucursal"
'-------------------------------------------------------------------------------------------'
Partial Class rToPagos_ConceptoSucursal

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            If lcParametro5Desde = 0 Then
                lcParametro5Desde = 7
            End If

            If Not (lcParametro6Desde = "") Then
                lcParametro6Desde = LTrim(RTrim(lcParametro6Desde))
                lcParametro6Desde = lcParametro6Desde.Replace(",,,,", "")
                lcParametro6Desde = lcParametro6Desde.Replace(",,,", "")
                lcParametro6Desde = lcParametro6Desde.Replace(",,", "")
                lcParametro6Desde = lcParametro6Desde.Replace(" ,  ", "")
                lcParametro6Desde = LTrim(RTrim(lcParametro6Desde))

                If lcParametro6Desde.Substring(lcParametro6Desde.Length - 1) = "," Then
                    lcParametro6Desde = lcParametro6Desde.Remove(lcParametro6Desde.Length - 1)
                End If

                If lcParametro6Desde.Substring(0) = "," Then
                    lcParametro6Desde = lcParametro6Desde.Remove(0)
                End If
            End If

            lcParametro6Desde = lcParametro6Desde.Replace("`", "'")
            lcParametro6Desde = lcParametro6Desde.Replace("´", "'")

            loComandoSeleccionar.AppendLine(" DECLARE Cur CURSOR FOR    ")
            If lcParametro6Desde = "" Then
                loComandoSeleccionar.AppendLine("     SELECT DISTINCT TOP " & lcParametro5Desde & " Nom_Suc FROM Sucursales   ")
            Else
                loComandoSeleccionar.AppendLine("     SELECT DISTINCT TOP " & lcParametro5Desde & " Nom_Suc FROM Sucursales WHERE Cod_Suc IN ('" & lcParametro6Desde.Replace(",", "','") & "')    ")
            End If
            loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine(" DECLARE		@Temp NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("             @Sucursales NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("             @SucursalesAux NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("             @SubConsulta NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("             @Numero INT    ")

            loComandoSeleccionar.AppendLine(" SET @SubConsulta = ''    ")
            loComandoSeleccionar.AppendLine(" SET @Sucursales = ''    ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux = ''    ")
            loComandoSeleccionar.AppendLine(" SET @Numero = 1    ")
            loComandoSeleccionar.AppendLine(" OPEN Cur    ")
            loComandoSeleccionar.AppendLine(" -- Consiguiendo todas las sucursales    ")
            loComandoSeleccionar.AppendLine("     FETCH NEXT FROM Cur INTO @Temp    ")
            loComandoSeleccionar.AppendLine("     WHILE @@FETCH_STATUS = 0    ")
            loComandoSeleccionar.AppendLine("             BEGIN    ")
            loComandoSeleccionar.AppendLine(" 		            SET @Sucursales = @Sucursales + '[' + RTrim(@Temp) + '],'    ")
            loComandoSeleccionar.AppendLine(" 		            SET @SucursalesAux = @SucursalesAux + 'ISNULL([' + RTrim(@Temp) + '],0) AS Columna'+CAST(@Numero AS NVARCHAR)+','    ")
            loComandoSeleccionar.AppendLine(" 		            SET @Numero = @Numero + 1    ")
            loComandoSeleccionar.AppendLine(" 		            FETCH NEXT FROM Cur INTO @Temp    ")

            loComandoSeleccionar.AppendLine("         End    ")
            loComandoSeleccionar.AppendLine(" CLOSE Cur    ")
            loComandoSeleccionar.AppendLine(" DEALLOCATE Cur    ")
            loComandoSeleccionar.AppendLine(" SET @Sucursales = SUBSTRING(@Sucursales, 0, LEN(@Sucursales))    ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux = SUBSTRING(@SucursalesAux, 0, LEN(@SucursalesAux))    ")
            loComandoSeleccionar.AppendLine(" -- Construyendo la consulta pivote    ")
            loComandoSeleccionar.AppendLine(" SET @SubConsulta = 'SELECT    ")
            loComandoSeleccionar.AppendLine(" 			            LTRIM(Sucursales.Nom_Suc) AS Nom_Suc,    ")
            loComandoSeleccionar.AppendLine(" 			            Conceptos.Nom_Con,     ")
            loComandoSeleccionar.AppendLine(" 			            Conceptos.Cod_Con,     ")
            loComandoSeleccionar.AppendLine(" 			            (Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_Hab) AS Saldo    ")
            loComandoSeleccionar.AppendLine("             INTO #Temp    ")
            loComandoSeleccionar.AppendLine("             FROM Ordenes_Pagos    ")
            loComandoSeleccionar.AppendLine("             JOIN Renglones_oPagos ON Renglones_oPagos.documento = Ordenes_Pagos.Documento     ")
            loComandoSeleccionar.AppendLine("             JOIN Conceptos ON Renglones_oPagos.Cod_Con = Conceptos.Cod_con    ")
            loComandoSeleccionar.AppendLine("             JOIN Sucursales ON Sucursales.Cod_Suc =  Ordenes_Pagos.cod_suc    ")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Fec_Ini Between '" & lcParametro0Desde & "'")
            loComandoSeleccionar.AppendLine("				AND '" & lcParametro0Hasta & "'")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Pro   Between '" & lcParametro1Desde & "'")
            loComandoSeleccionar.AppendLine("				AND '" & lcParametro1Hasta & "'")
            loComandoSeleccionar.AppendLine("				AND Renglones_oPagos.Cod_Con  Between '" & lcParametro2Desde & "'")
            loComandoSeleccionar.AppendLine("				AND '" & lcParametro2Hasta & "'")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Cod_Mon  Between '" & lcParametro3Desde & "'")
            loComandoSeleccionar.AppendLine("				AND '" & lcParametro3Hasta & "'")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Pagos.Status    IN (" & lcParametro4Desde.Replace("'", "''") & ")")

            loComandoSeleccionar.AppendLine(" SELECT Nom_Con AS Concepto, Cod_Con, ' + @SucursalesAux + '    ")
            loComandoSeleccionar.AppendLine(" FROM    ")
            loComandoSeleccionar.AppendLine("     (		    ")
            loComandoSeleccionar.AppendLine("             SELECT     ")
            loComandoSeleccionar.AppendLine(" 		            RTRIM(#Temp.Nom_Suc) AS Nom_Suc,    ")
            loComandoSeleccionar.AppendLine(" 		            #Temp.Nom_Con,    ")
            loComandoSeleccionar.AppendLine(" 		            #Temp.Cod_Con,    ")
            loComandoSeleccionar.AppendLine(" 		            #Temp.Saldo AS Saldo    ")
            loComandoSeleccionar.AppendLine("             FROM #Temp    ")
            loComandoSeleccionar.AppendLine("     ) AS SourceTable    ")
            loComandoSeleccionar.AppendLine(" PIVOT    ")
            loComandoSeleccionar.AppendLine("     (    ")
            loComandoSeleccionar.AppendLine(" SUM(Saldo)    ")
            loComandoSeleccionar.AppendLine("         FOR Nom_Suc IN (' + @Sucursales + ')    ")
            loComandoSeleccionar.AppendLine("     ) AS PivotTable    ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & ";' ")
            loComandoSeleccionar.AppendLine(" EXEC sp_executesql  @SubConsulta    ")

            If lcParametro6Desde = "" Then
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro5Desde & " RTRIM(Nom_Suc) AS Nom_Suc, RTRIM(Cod_Suc) AS Cod_Suc FROM Sucursales    ")
            Else
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro5Desde & " RTRIM(Nom_Suc) AS Nom_Suc, RTRIM(Cod_Suc) AS Cod_Suc FROM Sucursales WHERE Cod_Suc IN ('" & lcParametro6Desde.Replace(",", "','") & "')    ")
            End If


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rToPagos_ConceptoSucursal", laDatosReporte)

            Dim Tope As Integer
            Tope = laDatosReporte.Tables(1).Rows.Count

            If Tope < lcParametro5Desde Then
                lcParametro5Desde = Tope
            End If

            Select Case lcParametro5Desde
                Case 1
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}"
                Case 2
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}"
                Case 3
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}"
                Case 4
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}"
                Case 5
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}"
                Case 6
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text15"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}"
                Case 7
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text15"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text14"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}"
            End Select


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrToPagos_ConceptoSucursal.ReportSource = loObjetoReporte

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
' CMS: 19/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
