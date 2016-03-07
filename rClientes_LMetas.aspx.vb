'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_LMetas"
'-------------------------------------------------------------------------------------------'
Partial Class rClientes_LMetas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lnParametro0Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(0)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            If lnParametro0Desde = 0 Then
                lnParametro0Desde = Date.Now().Year()
            End If

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(lnParametro0Desde)
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    YEAR(Facturas.Fec_Ini)              AS  Año, ")
            loComandoSeleccionar.AppendLine("           MONTH(Facturas.Fec_Ini)             AS  Mes_Ini, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Cli                    AS  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Facturas.Mon_Net)     AS  Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Facturas.Can_Art1)    AS  Can_Art  ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpFacturas ")
            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento          =   Renglones_Facturas.Documento    ")
            loComandoSeleccionar.AppendLine("           AND YEAR(Facturas.Fec_Ini)  =   " & lnParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli        BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven        BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Status         In (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Tra        BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Mon        BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  YEAR(Facturas.Fec_Ini), MONTH(Facturas.Fec_Ini), Facturas.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Facturas.Cod_Cli, YEAR(Facturas.Fec_Ini), MONTH(Facturas.Fec_Ini) ")



            loComandoSeleccionar.AppendLine(" SELECT    Metas.Año           AS  Año, ")
            loComandoSeleccionar.AppendLine("           Metas.Mes           AS  Mes_Ini, ")
            loComandoSeleccionar.AppendLine("           Metas.Cod_Cli       AS  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli    AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Metas.Monto1        AS  Monto1, ")
            loComandoSeleccionar.AppendLine("           Metas.Can_Art1      AS  Can_Art1 ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpMetas ")
            loComandoSeleccionar.AppendLine(" FROM      Metas, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Metas.Cod_Cli       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Metas.Tip_Met   =   'Cliente_Mensualmente' ")
            loComandoSeleccionar.AppendLine("           AND Metas.Año       =   " & lnParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND Metas.Cod_Cli   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento & ",")
            loComandoSeleccionar.AppendLine("           Metas.Año, ")
            loComandoSeleccionar.AppendLine("           Metas.Mes ")


            loComandoSeleccionar.AppendLine(" SELECT    #tmpMetas.Año, ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Mes_Ini, ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Monto1, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpFacturas.Mon_Net, 0) AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpFacturas.Can_Art, 0) AS Can_Art ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpValores ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpMetas LEFT JOIN #tmpFacturas ON ")
            loComandoSeleccionar.AppendLine("           (#tmpMetas.Año    =   #tmpFacturas.Año       AND ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Mes_Ini =   #tmpFacturas.Mes_Ini   AND ")
            loComandoSeleccionar.AppendLine("           #tmpMetas.Cod_Cli =   #tmpFacturas.Cod_Cli)  ")

            loComandoSeleccionar.AppendLine(" SELECT    #tmpValores.Año, ")
            loComandoSeleccionar.AppendLine("   		CASE    WHEN #tmpValores.Mes_Ini=1  THEN 'Ene' ")
            loComandoSeleccionar.AppendLine("   		        WHEN #tmpValores.Mes_Ini=2  THEN 'Feb' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=3  THEN 'Mar' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=4  THEN 'Abr' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=5  THEN 'May' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=6  THEN 'Jun' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=7  THEN 'Jul' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=8  THEN 'Ago' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=9  THEN 'Sep' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=10 THEN 'Oct' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=11 THEN 'Nov' ")
            loComandoSeleccionar.AppendLine("			        WHEN #tmpValores.Mes_Ini=12 THEN 'Dic' ")
            loComandoSeleccionar.AppendLine("   		END                                                 AS  Mes, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Monto1, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           CASE    WHEN (#tmpValores.Monto1 = 0.00 AND #tmpValores.Mon_Net = 0.00) THEN 0.00 ")
            loComandoSeleccionar.AppendLine("                   WHEN (#tmpValores.Monto1 > 0.00 AND #tmpValores.Mon_Net = 0.00) THEN 0.00 ")
            loComandoSeleccionar.AppendLine("                   WHEN (#tmpValores.Monto1 = 0.00 AND #tmpValores.Mon_Net > 0.00) THEN 100 ")
            loComandoSeleccionar.AppendLine("           ELSE    ((#tmpValores.Mon_Net * 100 / #tmpValores.Monto1)) ")
            loComandoSeleccionar.AppendLine("           END  AS  Por_Mon, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           #tmpValores.Can_Art, ")
            loComandoSeleccionar.AppendLine("           CASE    WHEN (#tmpValores.Can_Art1 = 0.00 AND #tmpValores.Can_Art = 0.00) THEN 0.00 ")
            loComandoSeleccionar.AppendLine("                   WHEN (#tmpValores.Can_Art1 > 0.00 AND #tmpValores.Can_Art = 0.00) THEN 0.00 ")
            loComandoSeleccionar.AppendLine("                   WHEN (#tmpValores.Can_Art1 = 0.00 AND #tmpValores.Can_Art > 0.00) THEN 100 ")
            loComandoSeleccionar.AppendLine("           ELSE    ((#tmpValores.Can_Art * 100 / #tmpValores.Can_Art1)) ")
            loComandoSeleccionar.AppendLine("           END  AS  Por_Can ")
            loComandoSeleccionar.AppendLine(" FROM     #tmpValores ")





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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_LMetas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrClientes_LMetas.ReportSource = loObjetoReporte

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
' JJD: 03/05/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
