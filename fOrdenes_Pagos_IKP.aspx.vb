'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOrdenes_Pagos_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fOrdenes_Pagos_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		Ordenes_Pagos.Cod_Pro,")
            loConsulta.AppendLine("       	    Ordenes_Pagos.Status,")
            loConsulta.AppendLine("			    Proveedores.Nom_Pro, ")
            loConsulta.AppendLine("             Proveedores.Rif									AS  Rif, ")
            loConsulta.AppendLine("             Proveedores.Nit, ")
            loConsulta.AppendLine("             SUBSTRING(Proveedores.Dir_Fis,1, 200)				AS  Dir_Fis, ")
            loConsulta.AppendLine("			    Proveedores.Telefonos, ")
            loConsulta.AppendLine("             Proveedores.Fax,")
            loConsulta.AppendLine("             Ordenes_Pagos.Documento,")
            loConsulta.AppendLine("             Ordenes_Pagos.Fec_Ini,")
            loConsulta.AppendLine("             Ordenes_Pagos.Fec_Fin,")
            loConsulta.AppendLine("             Ordenes_Pagos.Mon_Bru,")
            loConsulta.AppendLine("             Ordenes_Pagos.Mon_Ret,")
            loConsulta.AppendLine("             Renglones_oPagos.Por_Imp1							AS  Por_Imp,")
            loConsulta.AppendLine("             CASE ")
            loConsulta.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1 ")
            loConsulta.AppendLine("            	ELSE ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Imp1 ")
            loConsulta.AppendLine("            END       As  Mon_Imp, ")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_ImpTot,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Net                              AS  Mon_Net,")
            loConsulta.AppendLine("            Ordenes_Pagos.Motivo                               AS  Motivo,")
            loConsulta.AppendLine("            Ordenes_Pagos.cod_suc                               AS  cod_suc,")
            loConsulta.AppendLine("            Renglones_oPagos.Renglon                           AS  Ren_Doc,")
            loConsulta.AppendLine("            Renglones_oPagos.Cod_Con                           AS  Cod_Con,")
            loConsulta.AppendLine("            Conceptos.Nom_Con		                            AS  Nom_Con,")
            loConsulta.AppendLine("            Renglones_oPagos.Cod_Imp                           AS  Cod_Imp,")
            loConsulta.AppendLine("            Renglones_oPagos.Mon_Deb                           AS  Mon_Deb,")
            loConsulta.AppendLine("            CASE ")
            loConsulta.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Net * -1 ")
            loConsulta.AppendLine("            	ELSE ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Net ")
            loConsulta.AppendLine("            END       As  NetoConcepto, ")            
            loConsulta.AppendLine("            Renglones_oPagos.Mon_Hab                           AS  Mon_Hab,")
            loConsulta.AppendLine(" 			 CONVERT(NCHAR(10), Ordenes_Pagos.Fec_Ini, 103)     AS	Fec_Che,")
            loConsulta.AppendLine("            0.00                                               AS  Ren_Tip,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Tip_Ope,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Doc_Des,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Num_Doc,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Num_Doc_Ban_Caj,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Cod_Caj,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Cod_Ban,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Cod_Cue,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Cod_Tar,")
            loConsulta.AppendLine("            0.00                                               AS  Mon_NetTP,")
            loConsulta.AppendLine("            'Conceptos'                                        AS  Tipo,")
            loConsulta.AppendLine("            '1'                                                AS  Orden,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Nom_Caj,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Nom_Ban,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Nom_Cue,")
            loConsulta.AppendLine("            SPACE(10)                                          AS  Nom_Tar")
            loConsulta.AppendLine("INTO		 #tmpDocumentos")
            loConsulta.AppendLine("FROM		 Ordenes_Pagos,")
            loConsulta.AppendLine("            Renglones_oPagos,")
            loConsulta.AppendLine("            Proveedores,")
            loConsulta.AppendLine("            conceptos")
            loConsulta.AppendLine("WHERE       Ordenes_Pagos.Documento    =   Renglones_oPagos.Documento AND ")
            loConsulta.AppendLine("            Ordenes_Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND")
            loConsulta.AppendLine("            Renglones_oPagos.Cod_Con   =   Conceptos.Cod_Con AND(" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	     Ordenes_Pagos.Cod_Pro,")
            loConsulta.AppendLine("            Ordenes_Pagos.Status, ")
            loConsulta.AppendLine("            Proveedores.Nom_Pro, ")
            loConsulta.AppendLine("            Proveedores.Rif AS  Rif, ")
            loConsulta.AppendLine("            Proveedores.Nit, ")
            loConsulta.AppendLine("            SUBSTRING(Proveedores.Dir_Fis,1, 200)  AS  Dir_Fis, ")
            loConsulta.AppendLine("            Proveedores.Telefonos AS  Telefonos, ")            
            loConsulta.AppendLine("            Proveedores.Fax,")
            loConsulta.AppendLine("            Ordenes_Pagos.Documento,")
            loConsulta.AppendLine("            Ordenes_Pagos.Fec_Ini,")
            loConsulta.AppendLine("            Ordenes_Pagos.Fec_Fin,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Bru,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Ret,")
            loConsulta.AppendLine("            0.00                                               AS  Por_Imp,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_Imp,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Imp                              AS  Mon_ImpTot,")
            loConsulta.AppendLine("            Ordenes_Pagos.Mon_Net                              AS  Mon_Net,")
            loConsulta.AppendLine("            Ordenes_Pagos.Motivo                               AS  Motivo,")
            loConsulta.AppendLine("            Ordenes_Pagos.cod_suc                               AS  cod_suc,")
            loConsulta.AppendLine(" 			0                                                  AS  Ren_Con,")
            loConsulta.AppendLine(" 			''                                                 AS  Cod_Con,")
            loConsulta.AppendLine(" 			''                                                 AS  Nom_Con,")
            loConsulta.AppendLine("			    ''                                                 AS  Cod_Imp,")
            loConsulta.AppendLine(" 		    0.00                                               AS  Mon_Deb,")
            loConsulta.AppendLine(" 		    0.00                                               AS  NetoConcepto,")
            loConsulta.AppendLine(" 		    0.00                                               AS  Mon_Hab,")
            loConsulta.AppendLine(" 			CONVERT(NCHAR(10), Detalles_oPagos.Fec_Ini, 103)   AS	Fec_Che,")
            loConsulta.AppendLine("            Detalles_oPagos.Renglon                            AS  Ren_Tip,")
            loConsulta.AppendLine("            Detalles_oPagos.Tip_Ope                            AS  Tip_Ope,")
            loConsulta.AppendLine("            Detalles_oPagos.Doc_Des                            AS  Doc_Des,")
            loConsulta.AppendLine("            Detalles_oPagos.Num_Doc                            AS  Num_Doc,")
            loConsulta.AppendLine("			 CASE")
            loConsulta.AppendLine("				WHEN Detalles_oPagos.Tip_Ope = 'Efectivo' Then Detalles_oPagos.Cod_Caj")
            loConsulta.AppendLine("				ELSE SUBSTRING(RTRIM(ISNULL((")
            loConsulta.AppendLine("					SELECT Bancos.Nom_Ban")
            loConsulta.AppendLine("					FROM Ordenes_Pagos ")
            loConsulta.AppendLine("				JOIN Detalles_oPagos ON (Ordenes_Pagos.Documento    =   Detalles_oPagos.Documento AND detalles_opagos.Renglon = '1')")
            loConsulta.AppendLine("				JOIN Cuentas_Bancarias ON (Cuentas_Bancarias.Cod_Cue = Detalles_oPagos.Cod_Cue AND detalles_opagos.Renglon = '1')")
            loConsulta.AppendLine("				JOIN Bancos ON (Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban)")
            loConsulta.AppendLine("				WHERE (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")
            loConsulta.AppendLine("				),'')),0,15)+': '+ Detalles_oPagos.Cod_Cue")
            loConsulta.AppendLine("				END AS  Num_Doc_Ban_Caj,")
            loConsulta.AppendLine("            Detalles_oPagos.Cod_Caj                            AS  Cod_Caj,")
            loConsulta.AppendLine("            Detalles_oPagos.Cod_Ban                            AS  Cod_Ban,")
            loConsulta.AppendLine("            Detalles_oPagos.Cod_Cue                            AS  Cod_Cue,")
            loConsulta.AppendLine("            Detalles_oPagos.Cod_Tar                            AS  Cod_Tar,")
            loConsulta.AppendLine("            Detalles_oPagos.Mon_Net                            AS  Mon_NetTP,")
            loConsulta.AppendLine("            'TiposPagos'                                       AS  Tipo, ")
            loConsulta.AppendLine("            '2'                                                AS  Orden ")
            loConsulta.AppendLine("INTO		#tmpTiposPagos1")
            loConsulta.AppendLine("FROM       Ordenes_Pagos, ")
            loConsulta.AppendLine("           Detalles_oPagos,")
            loConsulta.AppendLine("           Proveedores  ")
            loConsulta.AppendLine("WHERE      Ordenes_Pagos.Documento    =   Detalles_oPagos.Documento AND detalles_opagos.Renglon = '1'  AND")
            loConsulta.AppendLine("           Ordenes_Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	   #tmpTiposPagos1.*,")
            loConsulta.AppendLine("          Cajas.Nom_Caj   AS  Nom_Caj")
            loConsulta.AppendLine("INTO      #tmpTiposPagos2")
            loConsulta.AppendLine("FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loConsulta.AppendLine("          ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	   #tmpTiposPagos2.*, ")
            loConsulta.AppendLine("          Bancos.Nom_Ban   AS  Nom_Ban ")
            loConsulta.AppendLine("INTO      #tmpTiposPagos3")
            loConsulta.AppendLine("FROM      #tmpTiposPagos2 LEFT JOIN Bancos")
            loConsulta.AppendLine("          ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	   #tmpTiposPagos3.*,")
            loConsulta.AppendLine("          Cuentas_Bancarias.Nom_Cue   AS  Nom_Cue ")
            loConsulta.AppendLine("INTO      #tmpTiposPagos4")
            loConsulta.AppendLine("FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias")
            loConsulta.AppendLine("          ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT    #tmpTiposPagos4.*,")
            loConsulta.AppendLine("          Tarjetas.Nom_Tar   AS  Nom_Tar")
            loConsulta.AppendLine("INTO      #tmpTiposPagos5")
            loConsulta.AppendLine("FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas")
            loConsulta.AppendLine("          ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  T.*, ")
            loConsulta.AppendLine("        Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,")
            loConsulta.AppendLine("        COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.direccion                     AS Direccion_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.telefonos                     AS Telefono_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.fax                           AS Fax_Empresa_Cliente")
            loConsulta.AppendLine("FROM(           ")
            loConsulta.AppendLine("        SELECT    * ")
            loConsulta.AppendLine("        FROM      #tmpDocumentos ")
            loConsulta.AppendLine("        UNION ALL ")
            loConsulta.AppendLine("        SELECT    * ")
            loConsulta.AppendLine("        FROM      #tmpTiposPagos5 ")
            loConsulta.AppendLine("    ) AS T")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = T.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("ORDER BY  32")
            loConsulta.AppendLine("")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos1")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos2")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos3")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos4")
            'loConsulta.AppendLine("DROP TABLE #tmpDocumentos")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos5")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Pagos_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)	
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfOrdenes_Pagos_IKP.ReportSource = loObjetoReporte

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
' JJD: 01/08/13: Programacion inicial
'-------------------------------------------------------------------------------------------'
