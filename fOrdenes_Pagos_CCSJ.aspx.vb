'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOrdenes_Pagos_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class fOrdenes_Pagos_CCSJ

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		 Ordenes_Pagos.Cod_Pro,")
            loComandoSeleccionar.AppendLine("       	 Ordenes_Pagos.Status,")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("            SUBSTRING(Proveedores.Dir_Fis,1, 200)  AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Movil, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Documento,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Fec_Ini,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Fec_Fin,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Bru,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Ret,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Por_Imp1							AS  Por_Imp,")
            loComandoSeleccionar.AppendLine("            CASE ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1 ")
            loComandoSeleccionar.AppendLine("            	ELSE ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("            END       As  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_ImpTot,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Net                              AS  Mon_Net,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Motivo                               AS  Motivo,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Renglon                           AS  Ren_Doc,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Cod_Con                           AS  Cod_Con,")
            loComandoSeleccionar.AppendLine("            Conceptos.Nom_Con		                            AS  Nom_Con,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Cod_Imp                           AS  Cod_Imp,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Mon_Deb                           AS  Mon_Deb,")
            loComandoSeleccionar.AppendLine("            CASE ")
            loComandoSeleccionar.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net * -1 ")
            loComandoSeleccionar.AppendLine("            	ELSE ")
            loComandoSeleccionar.AppendLine("            		Renglones_oPagos.Mon_Net ")
            loComandoSeleccionar.AppendLine("            END       As  NetoConcepto, ")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Mon_Hab                           AS  Mon_Hab,")
            loComandoSeleccionar.AppendLine(" 			 CONVERT(NCHAR(10), Ordenes_Pagos.Fec_Ini, 103)     AS	Fec_Che,")
            loComandoSeleccionar.AppendLine("            0.00                                               AS  Ren_Tip,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Tip_Ope,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Doc_Des,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Num_Doc,")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Tip_Cue, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Num_Cue, ")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Num_Doc_Ban_Caj,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Cod_Caj,")
            loComandoSeleccionar.AppendLine("           Cuentas_Clientes.Cod_Ban    AS  Cod_Ban_Pro, ")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Cod_Ban,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Cod_Cue,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Cod_Tar,")
            loComandoSeleccionar.AppendLine("            0.00                                               AS  Mon_NetTP,")
            loComandoSeleccionar.AppendLine("            'Conceptos'                                        AS  Tipo,")
            loComandoSeleccionar.AppendLine("            '1'                                                AS  Orden,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Nom_Caj,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Nom_Ban,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Nom_Cue,")
            loComandoSeleccionar.AppendLine("            SPACE(10)                                          AS  Nom_Tar")
            loComandoSeleccionar.AppendLine("INTO		 #tmpDocumentos")
            loComandoSeleccionar.AppendLine("FROM		 Ordenes_Pagos,")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos,")
            loComandoSeleccionar.AppendLine("            conceptos,")
            loComandoSeleccionar.AppendLine("           Proveedores LEFT JOIN Cuentas_Clientes ")
            loComandoSeleccionar.AppendLine("           ON  Proveedores.Cod_Pro =   Cuentas_Clientes.Cod_Reg ")
            loComandoSeleccionar.AppendLine("WHERE       Ordenes_Pagos.Documento    =   Renglones_oPagos.Documento AND ")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos.Cod_Con   =   Conceptos.Cod_Con AND(" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loComandoSeleccionar.AppendLine("SELECT	     Ordenes_Pagos.Cod_Pro,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("            SUBSTRING(Proveedores.Dir_Fis,1, 200)  AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Movil, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Documento,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Fec_Ini,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Fec_Fin,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Bru,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Ret,")
            loComandoSeleccionar.AppendLine("            0.00                                               AS  Por_Imp,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_Imp,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_ImpTot,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Mon_Net                              AS  Mon_Net,")
            loComandoSeleccionar.AppendLine("            Ordenes_Pagos.Motivo                               AS  Motivo,")
            loComandoSeleccionar.AppendLine(" 			 0                                                  AS  Ren_Con,")
            loComandoSeleccionar.AppendLine(" 			 ''                                                 AS  Cod_Con,")
            loComandoSeleccionar.AppendLine(" 			 ''                                                 AS  Nom_Con,")
            loComandoSeleccionar.AppendLine("			 ''                                                 AS  Cod_Imp,")
            loComandoSeleccionar.AppendLine(" 			 0.00                                               AS  Mon_Deb,")
            loComandoSeleccionar.AppendLine(" 			 0.00                                               AS  NetoConcepto,")
            loComandoSeleccionar.AppendLine(" 			 0.00                                               AS  Mon_Hab,")
            loComandoSeleccionar.AppendLine(" 			 CONVERT(NCHAR(10), Detalles_oPagos.Fec_Ini, 103)	AS	Fec_Che,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Renglon                            AS  Ren_Tip,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Tip_Ope                            AS  Tip_Ope,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Doc_Des                            AS  Doc_Des,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Num_Doc                            AS  Num_Doc,")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Tip_Cue, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Num_Cue, ")
            loComandoSeleccionar.AppendLine("			 CASE")
            loComandoSeleccionar.AppendLine("				WHEN Detalles_oPagos.Tip_Ope = 'Efectivo' Then Detalles_oPagos.Cod_Caj")
            loComandoSeleccionar.AppendLine("				ELSE SUBSTRING(RTRIM(ISNULL((")
            loComandoSeleccionar.AppendLine("					SELECT Bancos.Nom_Ban")
            loComandoSeleccionar.AppendLine("					FROM Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("				JOIN Detalles_oPagos ON (Ordenes_Pagos.Documento    =   Detalles_oPagos.Documento AND detalles_opagos.Renglon = '1')")
            loComandoSeleccionar.AppendLine("				JOIN Cuentas_Bancarias ON (Cuentas_Bancarias.Cod_Cue = Detalles_oPagos.Cod_Cue AND detalles_opagos.Renglon = '1')")
            loComandoSeleccionar.AppendLine("				JOIN Bancos ON (Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban)")
            loComandoSeleccionar.AppendLine("				WHERE (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")
            loComandoSeleccionar.AppendLine("				),'')),0,15)+': '+ Detalles_oPagos.Cod_Cue")
            loComandoSeleccionar.AppendLine("				END AS  Num_Doc_Ban_Caj,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Cod_Caj                            AS  Cod_Caj,")
            loComandoSeleccionar.AppendLine("           Cuentas_Clientes.Cod_Ban    AS  Cod_Ban_Pro, ")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Cod_Ban                            AS  Cod_Ban,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Cod_Cue                            AS  Cod_Cue,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Cod_Tar                            AS  Cod_Tar,")
            loComandoSeleccionar.AppendLine("            Detalles_oPagos.Mon_Net                            AS  Mon_NetTP,")
            loComandoSeleccionar.AppendLine("            'TiposPagos'                                       AS  Tipo, ")
            loComandoSeleccionar.AppendLine("            '2'                                                AS  Orden ")
            loComandoSeleccionar.AppendLine("INTO		#tmpTiposPagos1")
            loComandoSeleccionar.AppendLine("FROM       Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("           Detalles_oPagos,")
            loComandoSeleccionar.AppendLine("           Proveedores LEFT JOIN Cuentas_Clientes ")
            loComandoSeleccionar.AppendLine("           ON  Proveedores.Cod_Pro =   Cuentas_Clientes.Cod_Reg ")
            loComandoSeleccionar.AppendLine("WHERE      Ordenes_Pagos.Documento    =   Detalles_oPagos.Documento AND detalles_opagos.Renglon = '1'  AND")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loComandoSeleccionar.AppendLine("SELECT	   #tmpTiposPagos1.*,")
            loComandoSeleccionar.AppendLine("          Cajas.Nom_Caj   AS  Nom_Caj")
            loComandoSeleccionar.AppendLine("INTO      #tmpTiposPagos2")
            loComandoSeleccionar.AppendLine("FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loComandoSeleccionar.AppendLine("          ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loComandoSeleccionar.AppendLine("SELECT	   #tmpTiposPagos2.*, ")
            loComandoSeleccionar.AppendLine("          Bancos.Nom_Ban   AS  Nom_Ban ")
            loComandoSeleccionar.AppendLine("INTO      #tmpTiposPagos3")
            loComandoSeleccionar.AppendLine("FROM      #tmpTiposPagos2 LEFT JOIN Bancos")
            loComandoSeleccionar.AppendLine("          ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")

            loComandoSeleccionar.AppendLine("SELECT	   #tmpTiposPagos3.*,")
            loComandoSeleccionar.AppendLine("          Cuentas_Bancarias.Nom_Cue   AS  Nom_Cue ")
            loComandoSeleccionar.AppendLine("INTO      #tmpTiposPagos4")
            loComandoSeleccionar.AppendLine("FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias")
            loComandoSeleccionar.AppendLine("          ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loComandoSeleccionar.AppendLine("SELECT    #tmpTiposPagos4.*,")
            loComandoSeleccionar.AppendLine("          Tarjetas.Nom_Tar   AS  Nom_Tar")
            loComandoSeleccionar.AppendLine("INTO      #tmpTiposPagos5")
            loComandoSeleccionar.AppendLine("FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas")
            loComandoSeleccionar.AppendLine("          ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loComandoSeleccionar.AppendLine("SELECT    *")
            loComandoSeleccionar.AppendLine("FROM      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT    * ")
            loComandoSeleccionar.AppendLine("FROM      #tmpTiposPagos5 ")
            loComandoSeleccionar.AppendLine("ORDER BY  32")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            'Me.mEscribirConsulta(loCOmandoSeleccionar.ToString())

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
      vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Pagos_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfOrdenes_Pagos_CCSJ.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' GCR: 19/03/09: Adicion de informacion de renglones y modificaciones al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 03/05/10: Se Ajusto para tomar el nombre del proveedor generico y que el debe sea (+)
'					y el haber (-)
'-------------------------------------------------------------------------------------------'
' MAT: 15/03/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'
' MAT: 30/04/11: Corrección para que muestre el Nombre del Beneficiario
'-------------------------------------------------------------------------------------------'
' JJD: 14/11/13: Programacion inicial del reporte
'-------------------------------------------------------------------------------------------'
