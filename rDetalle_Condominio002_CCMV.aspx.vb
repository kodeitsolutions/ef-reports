'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDetalle_Condominio002_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rDetalle_Condominio002_CCMV

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
            



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" Renglones_oPagos .*, Ordenes_Pagos.Fec_Ini As Fecha, Conceptos.Nom_Con")
            loComandoSeleccionar.AppendLine(" INTO #tempOrdenes")
            loComandoSeleccionar.AppendLine(" FROM Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_oPagos on Renglones_oPagos.Documento = Ordenes_pagos.Documento")
            loComandoSeleccionar.AppendLine(" Join Conceptos on Conceptos.Cod_Con = Renglones_oPagos.Cod_Con")
            loComandoSeleccionar.AppendLine(" WHERE Cod_Rev = '02'")
            loComandoSeleccionar.AppendLine("       And Ordenes_Pagos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       And Conceptos.Cod_Con <> 'E030'")

            'loComandoSeleccionar.AppendLine(" SELECT ")
            'loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Cla,")
            'loComandoSeleccionar.AppendLine(" 		Clientes.Cod_Cli,")
            'loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            'loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali,")
            'loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali2,")
            'loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali3")
            'loComandoSeleccionar.AppendLine(" INTO #tempClientesArticulos")
            'loComandoSeleccionar.AppendLine(" FROM Clientes")
            'loComandoSeleccionar.AppendLine(" JOIN Articulos On Articulos.Cod_Art = Clientes.Cod_Cli ")
            'loComandoSeleccionar.AppendLine(" WHERE Clientes.Status = 'A' And ")
            'loComandoSeleccionar.AppendLine(" 		Articulos.Status = 'A'")

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Cla,")
            loComandoSeleccionar.AppendLine(" 		    Precios_Clientes.Cod_Art AS Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		    Precios_Clientes.Cod_Reg AS Cod_Ori,")
            loComandoSeleccionar.AppendLine(" 		    Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		    Articulos.Por_Ali,")
            loComandoSeleccionar.AppendLine(" 		    Articulos.Por_Ali2,")
            loComandoSeleccionar.AppendLine(" 		    Articulos.Por_Ali3")
            loComandoSeleccionar.AppendLine(" INTO      #tempClientesArticulos")
            loComandoSeleccionar.AppendLine(" FROM      Articulos ")
            loComandoSeleccionar.AppendLine("           JOIN Precios_Clientes	ON	Articulos.Cod_Art = Precios_Clientes.Cod_Art ")
            loComandoSeleccionar.AppendLine("           JOIN Clientes			ON	Precios_Clientes.Cod_Reg = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" WHERE     Clientes.Status = 'A' And ")
            loComandoSeleccionar.AppendLine(" 		    Articulos.Status = 'A'")

            loComandoSeleccionar.AppendLine(" SELECT    #tempClientesArticulos.Cod_Ori, ")
            loComandoSeleccionar.AppendLine("           #tempClientesArticulos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 	        #tempClientesArticulos.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 	        #tempClientesArticulos.Por_Ali,")
            loComandoSeleccionar.AppendLine(" 	        #tempClientesArticulos.Por_Ali2,")
            loComandoSeleccionar.AppendLine(" 	        #tempClientesArticulos.Por_Ali3,")
            loComandoSeleccionar.AppendLine(" 	        SUM(#tempOrdenes.Mon_Net)                   As Mon_Net,")
            loComandoSeleccionar.AppendLine(" 	        CAST(#tempOrdenes.Nom_Con As Varchar(5000)) As Motivo,")
            loComandoSeleccionar.AppendLine("           DATEPART(YEAR, #tempOrdenes.Fecha)          As Anio, ")
            loComandoSeleccionar.AppendLine(" 	        DATEPART(MONTH, #tempOrdenes.Fecha)         As Mes,")
            loComandoSeleccionar.AppendLine(" 	        #tempOrdenes.Cod_Con,")
            loComandoSeleccionar.AppendLine("           SUM(Case When #tempOrdenes.Cod_Con In ('E024','E025') Then ")
            loComandoSeleccionar.AppendLine("           		#tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali2/100)")
            loComandoSeleccionar.AppendLine("           	Else")
            loComandoSeleccionar.AppendLine("           		#tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali/100)")
            loComandoSeleccionar.AppendLine("               End)                                    As NetxAli,")
            loComandoSeleccionar.AppendLine("           SUM(Case When #tempOrdenes.Cod_Con In ('E024','E025') Then ")
            loComandoSeleccionar.AppendLine("           		(#tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali2/100)) * 1 ")
            loComandoSeleccionar.AppendLine("           	Else")
            loComandoSeleccionar.AppendLine("           		(#tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali/100)) * 0.03")
            loComandoSeleccionar.AppendLine("               End)                                    As NetxAliPlus,")
            loComandoSeleccionar.AppendLine("           Case When #tempOrdenes.Cod_Con In ('E024','E025') Then ")
            loComandoSeleccionar.AppendLine("           		2")
            loComandoSeleccionar.AppendLine("           Else")
            loComandoSeleccionar.AppendLine("           		1")
            loComandoSeleccionar.AppendLine("           End                                         As Alicuota")
            loComandoSeleccionar.AppendLine("           ")
            loComandoSeleccionar.AppendLine(" INTO      #temFinal")
            loComandoSeleccionar.AppendLine(" FROM      #tempClientesArticulos, #tempOrdenes")

            loComandoSeleccionar.AppendLine(" WHERE     #tempOrdenes.Cod_Con Not In ('E024','E025')")
            loComandoSeleccionar.AppendLine("           And Cod_Cli                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And #tempOrdenes.Fecha     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And #tempClientesArticulos.Cod_Cla     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)



            loComandoSeleccionar.AppendLine(" GROUP BY  DATEPART(YEAR, #tempOrdenes.Fecha), ")
            loComandoSeleccionar.AppendLine(" 		    DATEPART(MONTH, #tempOrdenes.Fecha),")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Cod_Ori, ")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Por_Ali,")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Por_Ali2,")
            loComandoSeleccionar.AppendLine(" 		    #tempClientesArticulos.	Por_Ali3,")
            loComandoSeleccionar.AppendLine(" 		    Cast(#tempOrdenes.			Nom_Con As varchar(5000)),")
            loComandoSeleccionar.AppendLine(" 		    #tempOrdenes.Cod_Con")
            loComandoSeleccionar.AppendLine(" ORDER BY  DATEPART(YEAR, #tempOrdenes.Fecha), DATEPART(MONTH, #tempOrdenes.Fecha), Cod_Ori, Cod_Cli, Nom_Cli, CAST(#tempOrdenes.Nom_Con As Varchar(5000))")


            
            loComandoSeleccionar.AppendLine(" SELECT    Cod_Cli + ")
            loComandoSeleccionar.AppendLine("       	Case ")
            loComandoSeleccionar.AppendLine("       		When dbo.Concatena(Cod_Ori) = '' Then Cod_Ori ")
            loComandoSeleccionar.AppendLine("       		Else dbo.Concatena(Cod_Ori) ")
            loComandoSeleccionar.AppendLine("       	End As Cod_Cli,")
            loComandoSeleccionar.AppendLine("       	Nom_Cli,")
            loComandoSeleccionar.AppendLine("       	Por_Ali,")
            loComandoSeleccionar.AppendLine("       	Por_Ali2,")
            loComandoSeleccionar.AppendLine("       	Por_Ali3,")
            loComandoSeleccionar.AppendLine("       	Mon_Net,")
            loComandoSeleccionar.AppendLine("       	UPPER(Motivo) As Motivo,")
            loComandoSeleccionar.AppendLine("       	Anio,")
            loComandoSeleccionar.AppendLine("       	Mes,")
            loComandoSeleccionar.AppendLine("       	Cod_Con,")
            loComandoSeleccionar.AppendLine("       	NetxAli,")
            loComandoSeleccionar.AppendLine("       	NetxAliPlus,")
            If lcParametro3Desde = "''" Then
                loComandoSeleccionar.AppendLine("       		(Select Por_Imp1 From Impuestos Where Cod_Imp = 'EXE') As Por_Imp1,")
            Else
                loComandoSeleccionar.AppendLine("       		(Select Por_Imp1 From Impuestos Where Cod_Imp = " & lcParametro3Desde & ") As Por_Imp1,")
            End If
            loComandoSeleccionar.AppendLine("       		Alicuota")
            loComandoSeleccionar.AppendLine("       From #temFinal  WHERE dbo.Concatena(Cod_Ori) <> ''  Order By Anio, Mes, Cod_Ori, Cod_Cli, Cod_Con")


            'Response.Clear()
            'Response.Write("<html><body><pre>" & vbNewLine)
            'Response.Write(loComandoSeleccionar.ToString)
            'Response.Write("</pre></body></html>")
            'Response.Flush()
            'Response.End()

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDetalle_Condominio002_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrDetalle_Condominio002_CCMV.ReportSource = loObjetoReporte

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
' CMS: 02/07/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' JJD: 20/07/10: Ajustes para generar un recibo por cada cliente asociado.
'-------------------------------------------------------------------------------------------'