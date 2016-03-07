'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rComisiones_Vendedores_Russo"
'-------------------------------------------------------------------------------------------'
Partial Class rComisiones_Vendedores_Russo

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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT   ")
            loComandoSeleccionar.AppendLine("             Clientes.Nom_Cli,   ")
            loComandoSeleccionar.AppendLine("             cuentas_cobrar.Cod_Cli,   ")
            loComandoSeleccionar.AppendLine("             Vendedores.Nom_Ven,   ")
            loComandoSeleccionar.AppendLine("             cuentas_cobrar.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("             cuentas_cobrar.Documento,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Fec_Ini AS Fec_Ini_Fac,   ")
            loComandoSeleccionar.AppendLine("             Cobros.Fec_Ini  AS Fec_Ini_Cob,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Mon_bru,   ")
            loComandoSeleccionar.AppendLine("             Detalles_Cobros.Tip_Ope,    ")
            loComandoSeleccionar.AppendLine("             Detalles_Cobros.Mon_Net AS Mon_Net_Det,   ")
            loComandoSeleccionar.AppendLine("             DATEDIFF(Day, cuentas_cobrar.Fec_Ini, Detalles_Cobros.Fec_Ini) AS Dias,   ")
            loComandoSeleccionar.AppendLine("             Detalles_Cobros.Fec_Ini AS Fec_Ini_Det,   ")
            loComandoSeleccionar.AppendLine("             ROUND(((cuentas_cobrar.Mon_Net/Cobros.Mon_Net)*100),2) AS Porcentaje_net,   ")
            loComandoSeleccionar.AppendLine("             ROUND(((cuentas_cobrar.Mon_Bas1/Cobros.Mon_Net)*100),2) AS Porcentaje_bas,   ")
            loComandoSeleccionar.AppendLine("             Cobros.Documento AS Doc_Cob,   ")
            loComandoSeleccionar.AppendLine("             Detalles_Cobros.Num_Doc AS Referencia ")
            loComandoSeleccionar.AppendLine(" INTO #tmpFacturas   ")
            loComandoSeleccionar.AppendLine(" FROM    Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros ON Renglones_Cobros.Cod_Tip = Cuentas_Cobrar.Cod_Tip AND Renglones_Cobros.Doc_Ori = Cuentas_Cobrar.Documento   ")
            loComandoSeleccionar.AppendLine(" JOIN Cobros ON (Renglones_Cobros.Documento = Cobros.Documento AND Cobros.Mon_Net <> 0)")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros on Detalles_Cobros.Documento = Cobros.Documento   ")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli   ")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven =  Cuentas_Cobrar.Cod_Ven   ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_ven        BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli        BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.cod_tip = 'Fact'   ")
            loComandoSeleccionar.AppendLine(" ORDER BY Cuentas_Cobrar.cod_Ven, Cuentas_Cobrar.Documento   ")
            loComandoSeleccionar.AppendLine("    ")
            loComandoSeleccionar.AppendLine(" SELECT	CASE    ")
            loComandoSeleccionar.AppendLine(" 	            WHEN Cuentas_Cobrar.Comentario like 'Devoluci%' THEN Cuentas_Cobrar.Mon_Bru   ")
            loComandoSeleccionar.AppendLine(" 	            ELSE '0'   ")
            loComandoSeleccionar.AppendLine("             END AS DvFA,   ")
            loComandoSeleccionar.AppendLine("             CASE    ")
            loComandoSeleccionar.AppendLine(" 	            WHEN Cuentas_Cobrar.Comentario like 'DeVC%' THEN Cuentas_Cobrar.Mon_Bru   ")
            loComandoSeleccionar.AppendLine(" 	            ELSE '0'   ")
            loComandoSeleccionar.AppendLine("             END AS DeVC,   ")
            loComandoSeleccionar.AppendLine("             CASE    ")
            loComandoSeleccionar.AppendLine(" 	            WHEN Cuentas_Cobrar.Comentario like 'DxPP%' THEN Cuentas_Cobrar.Mon_Bru   ")
            loComandoSeleccionar.AppendLine(" 	            ELSE '0'   ")
            loComandoSeleccionar.AppendLine("             END AS DxPP,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Cod_Cli,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Doc_Ori,   ")
            loComandoSeleccionar.AppendLine("             Cuentas_Cobrar.Tip_Ori,   ")
            loComandoSeleccionar.AppendLine("             Renglones_Cobros.Doc_Ori AS Documento   ")
            loComandoSeleccionar.AppendLine(" INTO #tmpNotasCreditos   ")
            loComandoSeleccionar.AppendLine(" FROM Cuentas_Cobrar   ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros  ON Renglones_Cobros.Cod_Tip = 'N/CR' AND Renglones_Cobros.Doc_Ori = Cuentas_Cobrar.Documento   ")
            loComandoSeleccionar.AppendLine(" AND  Renglones_Cobros.Cod_Tip = Cuentas_Cobrar.Cod_tip   ")
            loComandoSeleccionar.AppendLine(" JOIN Cobros ON Renglones_Cobros.Documento = Cobros.Documento   ")
            loComandoSeleccionar.AppendLine(" where ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_ven        BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     AND Cuentas_Cobrar.Doc_Ori NOT IN ('')   ")
            loComandoSeleccionar.AppendLine("     AND Cuentas_Cobrar.Tip_Ori NOT IN ('')   ")
            'loComandoSeleccionar.AppendLine(" ORDER BY Cuentas_Cobrar.cod_Ven, Cuentas_Cobrar.Documento   ")

            loComandoSeleccionar.AppendLine("UNION ALL  ")

            loComandoSeleccionar.AppendLine("     SELECT	CASE      ")
            loComandoSeleccionar.AppendLine("                    WHEN Cuentas_Cobrar.Comentario like 'Devoluci%' THEN Cuentas_Cobrar.Mon_Bru     ")
            loComandoSeleccionar.AppendLine("                    ELSE '0'     ")
            loComandoSeleccionar.AppendLine("                 END AS DvFA,     ")
            loComandoSeleccionar.AppendLine("                 CASE      ")
            loComandoSeleccionar.AppendLine("                    WHEN Cuentas_Cobrar.Comentario like 'DeVC%' THEN Cuentas_Cobrar.Mon_Bru     ")
            loComandoSeleccionar.AppendLine("                    ELSE '0'     ")
            loComandoSeleccionar.AppendLine("                 END AS DeVC,     ")
            loComandoSeleccionar.AppendLine("                 CASE      ")
            loComandoSeleccionar.AppendLine("                    WHEN Cuentas_Cobrar.Comentario like 'DxPP%' THEN Cuentas_Cobrar.Mon_Bru     ")
            loComandoSeleccionar.AppendLine("                    ELSE '0'     ")
            loComandoSeleccionar.AppendLine("                 END AS DxPP,     ")
            loComandoSeleccionar.AppendLine("                 Cuentas_Cobrar.Cod_Cli,     ")
            loComandoSeleccionar.AppendLine("                 Cuentas_Cobrar.Cod_Ven,     ")
            loComandoSeleccionar.AppendLine("                 Cuentas_Cobrar.Factura AS Doc_Ori,     ")
            loComandoSeleccionar.AppendLine("                 'Fact' AS Tip_Ori,     ")
            loComandoSeleccionar.AppendLine("        Cuentas_Cobrar.Documento  ")
            loComandoSeleccionar.AppendLine("        FROM Cuentas_Cobrar  ")
            loComandoSeleccionar.AppendLine("      JOIN Cobros ON Cobros.Documento = Cuentas_Cobrar.Doc_Ori AND Cuentas_Cobrar.Tip_Ori = 'Cobros'     ")
            loComandoSeleccionar.AppendLine("     where Cuentas_Cobrar.Cod_Tip = 'N/CR'  ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_ven        BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     AND Cuentas_Cobrar.Doc_Ori NOT IN ('')   ")
            loComandoSeleccionar.AppendLine("     AND Cuentas_Cobrar.Tip_Ori NOT IN ('')   ")
           
            loComandoSeleccionar.AppendLine("    ")
            loComandoSeleccionar.AppendLine(" SELECT   ")
            loComandoSeleccionar.AppendLine("             ROW_NUMBER() OVER(PARTITION BY #tmpFacturas.Doc_Cob, #tmpFacturas.Documento ORDER BY #tmpFacturas.Cod_Ven ASC , #tmpFacturas.Doc_Cob, #tmpFacturas.Documento) AS Fila, ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Nom_Cli,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Nom_Ven,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Documento,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Fec_Ini_Fac,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Fec_Ini_Cob,   ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Mon_Bru AS Mon_Bru_Fac,   ")
            loComandoSeleccionar.AppendLine(" 	          isnull(#tmpNotasCreditos.DvFA, '0') AS DvFA, ")
            loComandoSeleccionar.AppendLine(" 	          isnull(#tmpNotasCreditos.DeVC, '0') AS DeVC,  ")
            loComandoSeleccionar.AppendLine(" 	          isnull(#tmpNotasCreditos.DxPP, '0') AS DxPP, ")
            loComandoSeleccionar.AppendLine(" 	          isnull(#tmpNotasCreditos.Doc_Ori, '0') AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Dias,   ")
            loComandoSeleccionar.AppendLine("             (CASE   ")
            loComandoSeleccionar.AppendLine(" 	           WHEN #tmpFacturas.Dias >= Parametros_Comisiones.Dia_Des AND #tmpFacturas.Dias <= Parametros_Comisiones.Dia_Has AND isnull(#tmpNotasCreditos.DxPP, '0') = 0  THEN  Parametros_Comisiones.Por_Com1    ")
            loComandoSeleccionar.AppendLine(" 	           WHEN #tmpFacturas.Dias >= Parametros_Comisiones.Dia_Des AND #tmpFacturas.Dias <= Parametros_Comisiones.Dia_Has AND isnull(#tmpNotasCreditos.DxPP, '0') > 0  THEN  Parametros_Comisiones.Por_Com2    ")
            loComandoSeleccionar.AppendLine(" 	           ELSE '0'      ")
            loComandoSeleccionar.AppendLine("             END) AS Por_Com,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Tip_Ope,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Mon_Net_Det,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Fec_Ini_Det,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Porcentaje_Net,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Porcentaje_Bas,    ")
            loComandoSeleccionar.AppendLine("             (#tmpFacturas.Porcentaje_Net*#tmpFacturas.Mon_Net_Det)/100 AS Net_Det,    ")
            loComandoSeleccionar.AppendLine("             (#tmpFacturas.Mon_Bru*(#tmpFacturas.Porcentaje_Net*#tmpFacturas.Mon_Net_Det))/100 AS Bas_Det,    ")
            loComandoSeleccionar.AppendLine("             #tmpFacturas.Doc_Cob,    ")
            loComandoSeleccionar.AppendLine("             CASE ")
            loComandoSeleccionar.AppendLine("                   WHEN #tmpFacturas.Tip_Ope = 'Efectivo' then 'Efectivo'")
            loComandoSeleccionar.AppendLine("                   ELSE #tmpFacturas.Referencia    ")
            loComandoSeleccionar.AppendLine("             END AS  Referencia")
            loComandoSeleccionar.AppendLine(" FROM	#tmpFacturas   ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN #tmpNotasCreditos ON  #tmpNotasCreditos.Doc_Ori = #tmpFacturas.Documento AND #tmpFacturas.Cod_Cli = #tmpNotasCreditos.Cod_Cli	AND #tmpFacturas.Cod_Ven = #tmpNotasCreditos.Cod_Ven   ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Parametros_Comisiones ON  Parametros_Comisiones.Cod_Ven = #tmpFacturas.Cod_Ven   ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & " , #tmpFacturas.Doc_Cob, #tmpFacturas.Documento")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComisiones_Vendedores_Russo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComisiones_Vendedores_Russo.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 28/02/09: Normalizacion del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 28/04/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'