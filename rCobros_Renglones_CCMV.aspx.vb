'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Renglones_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Renglones_CCMV

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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loConsulta.AppendLine("           Clientes.Nom_Cli, ")
            loConsulta.AppendLine("           Clientes.Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           Clientes.Dir_Fis, ")
            loConsulta.AppendLine("           Clientes.Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cobros.Documento, ")
            loConsulta.AppendLine("           Cobros.Fec_Ini, ")
            loConsulta.AppendLine("           Cobros.Fec_Fin, ")
            loConsulta.AppendLine("           Cobros.Mon_Bru, ")
            loConsulta.AppendLine("           Cobros.Mon_Imp, ")
            loConsulta.AppendLine("           Cobros.Mon_Net, ")
            loConsulta.AppendLine("           Cobros.Comentario, ")
            loConsulta.AppendLine("           Renglones_Cobros.Renglon    AS  Ren_Doc, ")
            loConsulta.AppendLine("           Renglones_Cobros.Tip_Doc    AS  Tip_Doc, ")
            loConsulta.AppendLine("           Renglones_Cobros.Cod_Tip    AS  Cod_Tip, ")
            loConsulta.AppendLine("           Renglones_Cobros.Doc_Ori    AS  Doc_Ori, ")
            loConsulta.AppendLine("           (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Net ELSE (Renglones_Cobros.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loConsulta.AppendLine("           (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Abo ELSE (Renglones_Cobros.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loConsulta.AppendLine("           0.00                        AS  Ren_Tip, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Tip_Ope, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Doc_Des, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Num_Doc, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Caj, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Ban, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Cue, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Tar, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_NetTP, ")
            loConsulta.AppendLine("           'Documentos'                AS  Tipo, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Caj, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Ban, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Cue, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Tar ")
            loConsulta.AppendLine(" INTO      #tmpDocumentos ")
            loConsulta.AppendLine(" FROM      Cobros, ")
            loConsulta.AppendLine("           Renglones_Cobros, ")
            loConsulta.AppendLine("           Cuentas_Cobrar, ")
            loConsulta.AppendLine("           Clientes ")
            loConsulta.AppendLine(" WHERE     Cobros.Documento                =   Renglones_Cobros.Documento ")
            loConsulta.AppendLine("           And Cobros.Cod_Cli              =   Clientes.Cod_Cli ")
            loConsulta.AppendLine("           And Cuentas_Cobrar.Cod_Tip      =   Renglones_Cobros.Cod_Tip ")
            loConsulta.AppendLine("           And Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori ")
            loConsulta.AppendLine("           And Cobros.Cod_Cli              =   Clientes.Cod_Cli ")
            loConsulta.AppendLine("           And Cobros.Documento            Between " & lcParametro0Desde)
            loConsulta.AppendLine("           And " & lcParametro0Hasta)
            loConsulta.AppendLine("           And Cobros.Fec_Ini              Between " & lcParametro1Desde)
            loConsulta.AppendLine("           And " & lcParametro1Hasta)
            loConsulta.AppendLine("           And Cobros.Cod_Cli              Between " & lcParametro2Desde)
            loConsulta.AppendLine("           And " & lcParametro2Hasta)
            loConsulta.AppendLine("           And Cobros.Status               IN (" & lcParametro3Desde & ")")
            loConsulta.AppendLine("           AND Cobros.Cod_Rev between " & lcParametro4Desde)
            loConsulta.AppendLine("    	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine("           AND Cobros.Cod_Suc between " & lcParametro5Desde)
            loConsulta.AppendLine("    	    AND " & lcParametro5Hasta)

            loConsulta.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loConsulta.AppendLine("           Clientes.Nom_Cli, ")
            loConsulta.AppendLine("           Clientes.Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           Clientes.Dir_Fis, ")
            loConsulta.AppendLine("           Clientes.Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cobros.Documento, ")
            loConsulta.AppendLine("           Cobros.Fec_Ini, ")
            loConsulta.AppendLine("           Cobros.Fec_Fin, ")
            loConsulta.AppendLine("           Cobros.Mon_Bru, ")
            loConsulta.AppendLine("           Cobros.Mon_Imp, ")
            loConsulta.AppendLine("           Cobros.Mon_Net, ")
            loConsulta.AppendLine("           Cobros.Comentario, ")
            loConsulta.AppendLine("           0                          AS  Ren_Doc, ")
            loConsulta.AppendLine("           ''                         AS  Tip_Doc, ")
            loConsulta.AppendLine("           ''                         AS  Cod_Tip, ")
            loConsulta.AppendLine("           ''                         AS  Doc_Ori, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_NetD, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_Abo, ")
            loConsulta.AppendLine("           Detalles_Cobros.Renglon    AS  Ren_Tip, ")
            loConsulta.AppendLine("           Detalles_Cobros.Tip_Ope    AS  Tip_Ope, ")
            loConsulta.AppendLine("           Detalles_Cobros.Doc_Des    AS  Doc_Des, ")
            loConsulta.AppendLine("           Detalles_Cobros.Num_Doc    AS  Num_Doc, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Caj    AS  Cod_Caj, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Ban    AS  Cod_Ban, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Cue    AS  Cod_Cue, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Tar    AS  Cod_Tar, ")
            loConsulta.AppendLine("           Detalles_Cobros.Mon_Net    AS  Mon_NetTP, ")
            loConsulta.AppendLine("           'TiposPagos'               AS  Tipo ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos1 ")
            loConsulta.AppendLine(" FROM      Cobros, ")
            loConsulta.AppendLine("           Detalles_Cobros, ")
            loConsulta.AppendLine("           Clientes ")
            loConsulta.AppendLine(" WHERE     Cobros.Documento        =   Detalles_Cobros.Documento ")
            loConsulta.AppendLine("           And Cobros.Cod_Cli      =   Clientes.Cod_Cli ")
            loConsulta.AppendLine("           And Cobros.Documento    Between " & lcParametro0Desde)
            loConsulta.AppendLine("           And " & lcParametro0Hasta)
            loConsulta.AppendLine("           And Cobros.Fec_Ini      Between " & lcParametro1Desde)
            loConsulta.AppendLine("           And " & lcParametro1Hasta)
            loConsulta.AppendLine("           And Cobros.Cod_Cli      Between " & lcParametro2Desde)
            loConsulta.AppendLine("           And " & lcParametro2Hasta)
            loConsulta.AppendLine("           And Cobros.Status       IN (" & lcParametro3Desde &")")
            loConsulta.AppendLine("           AND Cobros.Cod_Suc between " & lcParametro5Desde)
            loConsulta.AppendLine("    	    AND " & lcParametro5Hasta)

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loConsulta.AppendLine("           SUBSTRING(Cajas.Nom_Caj,1,25)   AS  Nom_Caj ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos2 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loConsulta.AppendLine("           SUBSTRING(Bancos.Nom_Ban,1,25)   AS  Nom_Ban ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos3 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loConsulta.AppendLine("           SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25)   AS  Nom_Cue ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos4 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loConsulta.AppendLine("           SUBSTRING(Tarjetas.Nom_Tar,1,25)   AS  Nom_Tar ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos5 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")
            loConsulta.AppendLine("")

            
            Dim llFiltrarCaja As Boolean = ((lcParametro6Desde<>"''") OrElse (lcParametro6Hasta<>"'zzzzzzz'"))
            Dim llFiltrarCuenta As Boolean = ((lcParametro7Desde<>"''") OrElse (lcParametro7Hasta<>"'zzzzzzz'"))
            
            If llFiltrarCaja OrElse llFiltrarCuenta Then 
                loConsulta.AppendLine("")
                loConsulta.AppendLine("CREATE TABLE #tmpFiltro(Documento CHAR(10) COLLATE DATABASE_DEFAULT);")
                loConsulta.AppendLine("")
                loConsulta.AppendLine("INSERT INTO #tmpFiltro(Documento)")
                loConsulta.AppendLine("SELECT  Documento")
                loConsulta.AppendLine("FROM    #tmpDocumentos D")
                loConsulta.AppendLine("WHERE   NOT EXISTS(SELECT * FROM #tmpTiposPagos5 TP WHERE TP.Documento = D.Documento)")
                loConsulta.AppendLine("UNION")
                loConsulta.AppendLine("SELECT  Documento")
                loConsulta.AppendLine("FROM    #tmpTiposPagos5")
                loConsulta.AppendLine("WHERE   tip_ope IN('Cheque','Efectivo','Tarjeta','Ticket') ")
                loConsulta.AppendLine("    AND Cod_Caj BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta & "")
                loConsulta.AppendLine("UNION")
                loConsulta.AppendLine("SELECT  Documento")
                loConsulta.AppendLine("FROM    #tmpTiposPagos5")
                loConsulta.AppendLine("WHERE   tip_ope IN('Deposito','Transferencia') ")
                loConsulta.AppendLine("    AND Cod_Cue BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & ";")
                loConsulta.AppendLine("")
                loConsulta.AppendLine("")
                loConsulta.AppendLine("DELETE FROM #tmpDocumentos")
                loConsulta.AppendLine("WHERE Documento NOT IN (SELECT Documento FROM #tmpFiltro);")
                loConsulta.AppendLine("")
                loConsulta.AppendLine("DELETE FROM #tmpTiposPagos5")
                loConsulta.AppendLine("WHERE Documento NOT IN (SELECT Documento FROM #tmpFiltro);")
                loConsulta.AppendLine("")
            End If

            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("SELECT        * ")
            loConsulta.AppendLine("FROM          #tmpDocumentos ")
            loConsulta.AppendLine("UNION ALL ")
            loConsulta.AppendLine("SELECT        * ")
            loConsulta.AppendLine("FROM          #tmpTiposPagos5 ") 
            loConsulta.AppendLine("ORDER BY       Documento, " & lcOrdenamiento)
            
            Dim loServicios As New cusDatos.goDatos()
            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCobros_Renglones_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Renglones_CCMV.ReportSource = loObjetoReporte

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
' JJD: 07/03/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' JJD: 14/03/09: Ajustes a los calculos de la los renglones de documentos cobrados			'
'-------------------------------------------------------------------------------------------'
' GCR: 30/03/09: Ajustes al diseño															'
'-------------------------------------------------------------------------------------------'
' CMS: 15/05/09: Filtro “Revisión:”														    '
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro “Sucursal:”														    '
'-------------------------------------------------------------------------------------------'
' CMS: 17/07/09: Metodo de Ordenamiento, Verificacion de registro							'
'-------------------------------------------------------------------------------------------'
' RJG: 10/04/12: Se agregó el total de documentos.											'
'-------------------------------------------------------------------------------------------'
