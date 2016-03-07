Imports System.Data
Partial Class fAnalisis_compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine("	SELECT 		")
            'loComandoSeleccionar.AppendLine("	ROW_NUMBER() OVER (ORDER BY Presupuestos.Cod_Pro) AS Renglon,		")
            'loComandoSeleccionar.AppendLine("	Presupuestos.cod_Pro, 		")
            'loComandoSeleccionar.AppendLine("	Proveedores.nom_Pro, 		")
            'loComandoSeleccionar.AppendLine("	Presupuestos.Documento 		")
            'loComandoSeleccionar.AppendLine("	INTO #temPresupuestos		")
            'loComandoSeleccionar.AppendLine("	From Presupuestos, Proveedores WHERE Presupuestos.Cod_Pro = Proveedores.Cod_Pro AND " & goFormatos.pcCondicionPrincipal)

            'loComandoSeleccionar.AppendLine("			")
            'loComandoSeleccionar.AppendLine("	  SELECT	Presupuestos.Documento,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Fec_Ini,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Fec_Fin,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Cod_Pro,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Cod_Ven,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Cod_Tra,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Status,   		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Renglon,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Cod_For,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Comentario,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Mon_Net As Mon_Net_Enc,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Mon_Bru As Mon_Bru_Enc,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Mon_Des1 As Mon_Des_Enc,  		")
            'loComandoSeleccionar.AppendLine("	 			Presupuestos.Dis_Imp,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Cod_Art,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Cod_Alm,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Precio1,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Can_Art1,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Cod_Uni,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Por_Imp1,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Por_Des,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos.Mon_Net,  		")
            'loComandoSeleccionar.AppendLine("	 			Articulos.Nom_Art,  		")
            'loComandoSeleccionar.AppendLine("	 			Vendedores.Nom_Ven,  		")
            'loComandoSeleccionar.AppendLine("	 			Transportes.Nom_Tra,  		")
            'loComandoSeleccionar.AppendLine("	 			Formas_Pagos.Nom_For,  		")
            'loComandoSeleccionar.AppendLine("	 			SUBSTRING(Proveedores.Nom_Pro,1,60) AS  Nom_Pro  		")
            'loComandoSeleccionar.AppendLine("	  INTO #temDatos		")
            'loComandoSeleccionar.AppendLine("	  FROM		Presupuestos,  		")
            'loComandoSeleccionar.AppendLine("	 			Renglones_Presupuestos,  		")
            'loComandoSeleccionar.AppendLine("	 			Proveedores,  		")
            'loComandoSeleccionar.AppendLine("	 			Articulos,  		")
            'loComandoSeleccionar.AppendLine("	 			Vendedores,  		")
            'loComandoSeleccionar.AppendLine("	 			Formas_Pagos,  		")
            'loComandoSeleccionar.AppendLine("	 			Transportes  		")
            'loComandoSeleccionar.AppendLine("	  WHERE		Presupuestos.Documento				=	Renglones_Presupuestos.Documento  		")
            'loComandoSeleccionar.AppendLine("	 			And Presupuestos.Cod_Pro			=	Proveedores.Cod_Pro  		")
            'loComandoSeleccionar.AppendLine("	 			And Renglones_Presupuestos.Cod_Art	=	Articulos.Cod_Art  		")
            'loComandoSeleccionar.AppendLine("	 			And Presupuestos.Cod_Ven			=	Vendedores.Cod_Ven  		")
            'loComandoSeleccionar.AppendLine("	 			And Presupuestos.Cod_For			=	Formas_Pagos.Cod_For  		")
            'loComandoSeleccionar.AppendLine("	 			And Presupuestos.Cod_Tra			=	Transportes.Cod_Tra  		")
            'loComandoSeleccionar.AppendLine("			And " & goFormatos.pcCondicionPrincipal)
            'loComandoSeleccionar.AppendLine("			")
            'loComandoSeleccionar.AppendLine("	SELECT 		")
            'loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Documento From #temPresupuestos Where #temPresupuestos.Renglon = 1)    Documento1,		")
            'loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Documento From #temPresupuestos Where #temPresupuestos.Renglon = 2)    Documento2,		")
            'loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Documento From #temPresupuestos Where #temPresupuestos.Renglon = 3)    Documento3,		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 1)     Cod_Pro1,		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 2)     Cod_Pro2,		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 3)     Cod_Pro3,		")
            'loComandoSeleccionar.AppendLine("	#temDatos.Cod_Art,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 1 Then #temDatos.Cod_Art Else '' End As Cod_Art1,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 2 Then #temDatos.Cod_Art Else '' End As Cod_Art2,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 3 Then #temDatos.Cod_Art Else '' End As Cod_Art3,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Precio1 Else '' End) As Precio11,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Precio1 Else '' End) As Precio12,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Precio1 Else '' End) As Precio13, 		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Can_Art1 Else '' End) As Can_Art11,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Can_Art1 Else '' End) As Can_Art12,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Can_Art1 Else '' End) As Can_Art13, 		")
            'loComandoSeleccionar.AppendLine("	#temDatos.Cod_Uni,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 1 Then #temDatos.Cod_Uni Else '' End As Cod_Uni1,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 2 Then #temDatos.Cod_Uni Else '' End As Cod_Uni2,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 3 Then #temDatos.Cod_Uni Else '' End As Cod_Uni3, 		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Por_Des Else '' End) As Por_Des1,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Por_Des Else '' End) As Por_Des2,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Por_Des Else '' End) As Por_Des3, 		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Mon_Net Else '' End) As Mon_Net1,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Mon_Net Else '' End) As Mon_Net2,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Mon_Net Else '' End) As Mon_Net3, 		")
            'loComandoSeleccionar.AppendLine("	#temDatos.Nom_Art ,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 1 Then #temDatos.Nom_Art Else '' End As Nom_Art1,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 2 Then #temDatos.Nom_Art Else '' End As Nom_Art2,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 3 Then #temDatos.Nom_Art Else '' End As Nom_Art3, 		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Nom_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 1) Nom_Pro1,		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Nom_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 2) Nom_Pro2,		")
            ''loComandoSeleccionar.AppendLine("	(Select Top 1 #temPresupuestos.Nom_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 3) Nom_Pro3,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 1 Then Cast(#temDatos.Comentario As varchar) Else '' End As Comentario1,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 2 Then Cast(#temDatos.Comentario As varchar) Else ''  End As Comentario2,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 3 Then Cast(#temDatos.Comentario As varchar) Else '' End As Comentario3,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 1 Then Cast(#temDatos.Dis_Imp As varchar) Else '' End As Dis_Imp1,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 2 Then Cast(#temDatos.Dis_Imp As varchar) Else ''  End As Dis_Imp2,		")
            'loComandoSeleccionar.AppendLine("	Case When #temPresupuestos.Renglon = 3 Then Cast(#temDatos.Dis_Imp As varchar) Else '' End As Dis_Imp3,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Mon_Net_Enc Else '' End) As Mon_Net_Enc1,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Mon_Net_Enc Else '' End) As Mon_Net_Enc2,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Mon_Net_Enc Else '' End) As Mon_Net_Enc3,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Mon_Bru_Enc Else '' End) As Mon_Bru_Enc1,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Mon_Bru_Enc Else '' End) As Mon_Bru_Enc2,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Mon_Bru_Enc Else '' End) As Mon_Bru_Enc3,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 1 Then #temDatos.Mon_Des_Enc Else '' End) As Mon_Des_Enc1,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 2 Then #temDatos.Mon_Des_Enc Else '' End) As Mon_Des_Enc2,		")
            'loComandoSeleccionar.AppendLine("	SUM(Case When #temPresupuestos.Renglon = 3 Then #temDatos.Mon_Des_Enc Else '' End) As Mon_Des_Enc3		")
            'loComandoSeleccionar.AppendLine("	INTO #tempFinal  FROM #temDatos 		")
            'loComandoSeleccionar.AppendLine("	JOIN #temPresupuestos On #temPresupuestos.Cod_Pro = #temDatos.Cod_Pro AND #temPresupuestos.Documento = #temDatos.Documento")

            'loComandoSeleccionar.AppendLine("	GROUP BY 	")
            ''loComandoSeleccionar.AppendLine("			#temDatos.cod_pro, 	")
            'loComandoSeleccionar.AppendLine("			#temDatos.cod_art, 	")
            'loComandoSeleccionar.AppendLine("			#temDatos.Documento, 	")
            'loComandoSeleccionar.AppendLine("			#temPresupuestos.Renglon, 	")
            'loComandoSeleccionar.AppendLine("			Cast(#temDatos.Dis_Imp As varchar), 	")
            'loComandoSeleccionar.AppendLine("			Cast(#temDatos.Comentario As varchar),	")
            'loComandoSeleccionar.AppendLine("			#temDatos.Nom_Art, 	")
            ''loComandoSeleccionar.AppendLine("			#temPresupuestos.Nom_Pro, 	")
            'loComandoSeleccionar.AppendLine("			#temDatos.Cod_Uni	")
            'loComandoSeleccionar.AppendLine("ORDER BY      #temDatos.Documento, #temDatos.cod_art")


            'loComandoSeleccionar.AppendLine("	Select  			")
            'loComandoSeleccionar.AppendLine("	IsNUll((Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 1)  , '') as   Cod_Pro1,				")
            'loComandoSeleccionar.AppendLine("	IsNUll((Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 2)  , '') as   Cod_Pro2,				")
            'loComandoSeleccionar.AppendLine("	IsNUll( (Select Top 1 #temPresupuestos.Cod_Pro From #temPresupuestos Where #temPresupuestos.Renglon = 3) , '') as    Cod_Pro3,				")
            'loComandoSeleccionar.AppendLine("	IsNUll(#tempFinal.Documento1, '')As Documento1, 		")
            'loComandoSeleccionar.AppendLine("	IsNUll(#tempFinal.Documento2, '') As Documento2,		")
            'loComandoSeleccionar.AppendLine("	IsNUll(#tempFinal.Documento3, '') As Documento3,		")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Art   ,			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Art1, 			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Art2, 			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Art3,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Precio11,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Precio12,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Precio13,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Can_Art11,     			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Can_Art12,     			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Can_Art13,     			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Uni,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Uni1,			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Uni2,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Cod_Uni3,   			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Por_Des1,     			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Por_Des2, 			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Por_Des3, 			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net1,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net2,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net3,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Nom_Art, 			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Nom_Art1,    			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Nom_Art2,   			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Nom_Art3,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Comentario1,   			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Comentario2,   			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Comentario3,   			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Dis_Imp1,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Dis_Imp2,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Dis_Imp3,      			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net_Enc1,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net_Enc2,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Net_Enc3,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Bru_Enc1,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Bru_Enc2,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Bru_Enc3,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Des_Enc1,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Des_Enc2,  			")
            'loComandoSeleccionar.AppendLine("	#tempFinal.Mon_Des_Enc3			")
            'loComandoSeleccionar.AppendLine("				")
            
           
            'loComandoSeleccionar.AppendLine("	 From #tempFinal			")


			loComandoSeleccionar.AppendLine("	SELECT 		 ")
			loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER (ORDER BY Presupuestos.Cod_Pro) AS Renglon,		 ")
			loComandoSeleccionar.AppendLine("		Presupuestos.cod_Pro, 		 ")
			loComandoSeleccionar.AppendLine("		Proveedores.nom_Pro, 		 ")
			loComandoSeleccionar.AppendLine("		Presupuestos.Documento 		 ")
			loComandoSeleccionar.AppendLine("	INTO #temPresupuestos		 ")
			loComandoSeleccionar.AppendLine("	From Presupuestos, Proveedores WHERE Presupuestos.Cod_Pro = Proveedores.Cod_Pro AND " & goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("	 ")
			loComandoSeleccionar.AppendLine("	 ")
			loComandoSeleccionar.AppendLine("	SELECT	Presupuestos.Documento,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Fec_Ini,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Fec_Fin,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Cod_Pro,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Cod_Ven,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Cod_Tra,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Status,   		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Renglon,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Cod_For,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Comentario,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Mon_Net As Mon_Net_Enc,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Mon_Bru As Mon_Bru_Enc,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Mon_Des1 As Mon_Des_Enc,  		 ")
			loComandoSeleccionar.AppendLine("		 			Presupuestos.Dis_Imp,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Cod_Art,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Cod_Alm,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Precio1,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Can_Art1,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Cod_Uni,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Por_Imp1,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Por_Des,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos.Mon_Net,  		 ")
			loComandoSeleccionar.AppendLine("		 			Articulos.Nom_Art,  		 ")
			loComandoSeleccionar.AppendLine("		 			Vendedores.Nom_Ven,  		 ")
			loComandoSeleccionar.AppendLine("		 			Transportes.Nom_Tra,  		 ")
			loComandoSeleccionar.AppendLine("		 			Formas_Pagos.Nom_For,  		 ")
			loComandoSeleccionar.AppendLine("		 			SUBSTRING(Proveedores.Nom_Pro,1,60) AS  Nom_Pro  		 ")
			loComandoSeleccionar.AppendLine("		    INTO #temDatos		 ")
			loComandoSeleccionar.AppendLine("			FROM		Presupuestos,  		 ")
			loComandoSeleccionar.AppendLine("		 			Renglones_Presupuestos,  		 ")
			loComandoSeleccionar.AppendLine("		 			Proveedores,  		 ")
			loComandoSeleccionar.AppendLine("		 			Articulos,  		 ")
			loComandoSeleccionar.AppendLine("		 			Vendedores,  		 ")
			loComandoSeleccionar.AppendLine("		 			Formas_Pagos,  		 ")
			loComandoSeleccionar.AppendLine("		 			Transportes  		 ")
			loComandoSeleccionar.AppendLine("		  WHERE		Presupuestos.Documento				=	Renglones_Presupuestos.Documento  		 ")
			loComandoSeleccionar.AppendLine("		 			And Presupuestos.Cod_Pro			=	Proveedores.Cod_Pro  		 ")
			loComandoSeleccionar.AppendLine("		 			And Renglones_Presupuestos.Cod_Art	=	Articulos.Cod_Art  		 ")
			loComandoSeleccionar.AppendLine("		 			And Presupuestos.Cod_Ven			=	Vendedores.Cod_Ven  		 ")
			loComandoSeleccionar.AppendLine("		 			And Presupuestos.Cod_For			=	Formas_Pagos.Cod_For  		 ")
			loComandoSeleccionar.AppendLine("		 			And Presupuestos.Cod_Tra			=	Transportes.Cod_Tra  		 ")
			loComandoSeleccionar.AppendLine("			And " & goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("		 ")
			loComandoSeleccionar.AppendLine("		 ")
			loComandoSeleccionar.AppendLine("		 SELECT	Distinct Articulos.Cod_Art As Cod_Art, Articulos.Nom_Art")
			loComandoSeleccionar.AppendLine("		 INTO #temArticulos		 ")
			loComandoSeleccionar.AppendLine("		 FROM		Presupuestos,  		 ")
			loComandoSeleccionar.AppendLine("		 		Renglones_Presupuestos,  		 ")
			loComandoSeleccionar.AppendLine("		 		Proveedores,  		 ")
			loComandoSeleccionar.AppendLine("		 		Articulos,  		 ")
			loComandoSeleccionar.AppendLine("		 		Vendedores,  		 ")
			loComandoSeleccionar.AppendLine("		 		Formas_Pagos,  		 ")
			loComandoSeleccionar.AppendLine("		 		Transportes  		 ")
			loComandoSeleccionar.AppendLine("		 WHERE		Presupuestos.Documento				=	Renglones_Presupuestos.Documento  ")
			loComandoSeleccionar.AppendLine("		 		And Presupuestos.Cod_Pro			=	Proveedores.Cod_Pro  		 ")
			loComandoSeleccionar.AppendLine("		 		And Renglones_Presupuestos.Cod_Art	=	Articulos.Cod_Art  		 ")
			loComandoSeleccionar.AppendLine("		 		And Presupuestos.Cod_Ven			=	Vendedores.Cod_Ven  		 ")
			loComandoSeleccionar.AppendLine("		 		And Presupuestos.Cod_For			=	Formas_Pagos.Cod_For  		 ")
			loComandoSeleccionar.AppendLine("		 		And Presupuestos.Cod_Tra			=	Transportes.Cod_Tra  		 ")
			loComandoSeleccionar.AppendLine("			And " & goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("		 ")
			loComandoSeleccionar.AppendLine("		 Select ")
			loComandoSeleccionar.AppendLine("		 Cod_Art, Nom_Art,")
			loComandoSeleccionar.AppendLine("		 ROW_NUMBER() OVER (ORDER BY Cod_Art) as ROW")
			loComandoSeleccionar.AppendLine("		 Into #temArticulos2")
			loComandoSeleccionar.AppendLine("		 from #temArticulos")
			loComandoSeleccionar.AppendLine("		 ")
			loComandoSeleccionar.AppendLine("		 ")
			loComandoSeleccionar.AppendLine("	SELECT 		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Documento As Documento1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Pro As Cod_Pro1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Pro As Nom_Pro1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art As Cod_Art1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Precio1 As Precio11,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Can_Art1 As Can_Art11,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni As Cod_Uni1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Por_Des As Por_Des1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net As Mon_Net1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art ,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art As Nom_Art1,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Comentario As varchar) As Comentario1,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Dis_Imp As varchar(6000)) As Dis_Imp1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net_Enc As Mon_Net_Enc1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Bru_Enc As Mon_Bru_Enc1,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Des_Enc As Mon_Des_Enc1		 ")
			loComandoSeleccionar.AppendLine("		INTO #tempPresupuestoProveedor1 ")
			loComandoSeleccionar.AppendLine("		FROM #temDatos 		 ")
			loComandoSeleccionar.AppendLine("		JOIN #temPresupuestos On #temPresupuestos.Cod_Pro = #temDatos.Cod_Pro AND #temPresupuestos.Documento = #temDatos.Documento ")
			loComandoSeleccionar.AppendLine("		Where #temPresupuestos.Renglon = 1 ")
			loComandoSeleccionar.AppendLine("		ORDER BY      #temDatos.Documento, #temDatos.cod_art ")
			loComandoSeleccionar.AppendLine("	 ")
			loComandoSeleccionar.AppendLine("	 ")
			loComandoSeleccionar.AppendLine("	SELECT 		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Documento As Documento2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Pro As Cod_Pro2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Pro As Nom_Pro2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art As Cod_Art2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Precio1 As Precio12,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Can_Art1 As Can_Art12,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni As Cod_Uni2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Por_Des As Por_Des2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net As Mon_Net2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art ,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art As Nom_Art2,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Comentario As varchar) As Comentario2,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Dis_Imp As varchar(6000)) As Dis_Imp2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net_Enc As Mon_Net_Enc2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Bru_Enc As Mon_Bru_Enc2,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Des_Enc As Mon_Des_Enc2		 ")
			loComandoSeleccionar.AppendLine("		INTO #tempPresupuestoProveedor2 ")
			loComandoSeleccionar.AppendLine("		FROM #temDatos 		 ")
			loComandoSeleccionar.AppendLine("		JOIN #temPresupuestos On #temPresupuestos.Cod_Pro = #temDatos.Cod_Pro AND #temPresupuestos.Documento = #temDatos.Documento ")
			loComandoSeleccionar.AppendLine("		Where #temPresupuestos.Renglon = 2 ")
			loComandoSeleccionar.AppendLine("		ORDER BY      #temDatos.Documento, #temDatos.cod_art ")
			loComandoSeleccionar.AppendLine("	 ")
			loComandoSeleccionar.AppendLine("	SELECT 		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Documento As Documento3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Pro As Cod_Pro3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Pro As Nom_Pro3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Art As Cod_Art3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Precio1 As Precio13,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Can_Art1 As Can_Art13,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Cod_Uni As Cod_Uni3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Por_Des As Por_Des3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net As Mon_Net3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art ,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Nom_Art As Nom_Art3,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Comentario As varchar) As Comentario3,		 ")
			loComandoSeleccionar.AppendLine("		Cast(#temDatos.Dis_Imp As varchar(6000)) As Dis_Imp3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Net_Enc As Mon_Net_Enc3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Bru_Enc As Mon_Bru_Enc3,		 ")
			loComandoSeleccionar.AppendLine("		#temDatos.Mon_Des_Enc As Mon_Des_Enc3		 ")
			loComandoSeleccionar.AppendLine("		INTO #tempPresupuestoProveedor3 ")
			loComandoSeleccionar.AppendLine("		FROM #temDatos 		 ")
			loComandoSeleccionar.AppendLine("		JOIN #temPresupuestos On #temPresupuestos.Cod_Pro = #temDatos.Cod_Pro AND #temPresupuestos.Documento = #temDatos.Documento ")
			loComandoSeleccionar.AppendLine("		Where #temPresupuestos.Renglon = 3 ")
			loComandoSeleccionar.AppendLine("		ORDER BY      #temDatos.Documento, #temDatos.cod_art ")
			
			
			'loComandoSeleccionar.AppendLine("	 SELECT ")
			'loComandoSeleccionar.AppendLine("			ROW_NUMBER() OVER (ORDER BY #tempPresupuestoProveedor1.Cod_Pro1) as ROW,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Documento1,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Cod_Pro1,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Nom_Pro1,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Cod_Art1,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Precio11,0) AS Precio11,   ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Can_Art11,0) AS Can_Art11,   ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Cod_Uni1,   ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Por_Des1,0) AS Por_Des1,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Mon_Net1,0) AS Mon_Net1,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Nom_Art1,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Comentario1, ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Dis_Imp1,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Mon_Net_Enc1,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Mon_Bru_Enc1,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor1.Mon_Des_Enc1,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Documento2, ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Cod_Pro2,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Nom_Pro2,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Cod_Art2,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor2.Precio12,0) As Precio12,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor2.Can_Art12,0) As Can_Art12,   ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Cod_Uni2,   ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor2.Por_Des2,0) As Por_Des2,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor2.Mon_Net2,0) As Mon_Net2,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Nom_Art2,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Comentario2, ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Dis_Imp2,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Mon_Net_Enc2,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Mon_Bru_Enc2,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor2.Mon_Des_Enc2,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Documento3, ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Cod_Pro3,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Nom_Pro3,     ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Cod_Art3,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor3.Precio13,0) As Precio13,    ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor3.Can_Art13,0) As Can_Art13,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Cod_Uni3,   ")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor3.Por_Des3,0) As Por_Des3,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor3.Mon_Net3,0) As Mon_Net3,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Nom_Art3,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Comentario3, ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Dis_Imp3,    ")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Mon_Net_Enc3,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Mon_Bru_Enc3,")
			'loComandoSeleccionar.AppendLine("	 		#tempPresupuestoProveedor3.Mon_Des_Enc3,")
			''loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Cod_Art, ISNULL(#tempPresupuestoProveedor2.Cod_Art, #tempPresupuestoProveedor3.Cod_Art)) As Cod_Art,")
			''loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Nom_Art, ISNULL(#tempPresupuestoProveedor2.Nom_Art, #tempPresupuestoProveedor3.Nom_Art)) As Nom_Art,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Cod_Uni, ISNULL(#tempPresupuestoProveedor2.Cod_Uni, #tempPresupuestoProveedor3.Cod_Uni)) As Cod_Uni,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Cod_Pro1, ISNULL(#tempPresupuestoProveedor2.Cod_Pro2, #tempPresupuestoProveedor3.Cod_Pro3)) As Cod_Pro,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Nom_Pro1, ISNULL(#tempPresupuestoProveedor2.Nom_Pro2, #tempPresupuestoProveedor3.Nom_Pro3)) As Nom_Pro,")
			'loComandoSeleccionar.AppendLine("	 		ISNULL(#tempPresupuestoProveedor1.Documento1, ISNULL(#tempPresupuestoProveedor2.Documento2, #tempPresupuestoProveedor3.Documento3)) As Documento")
			'loComandoSeleccionar.AppendLine("	 INTO #tempFinal ")
			'loComandoSeleccionar.AppendLine("	 FROM ")
			'loComandoSeleccionar.AppendLine("	 #tempPresupuestoProveedor1 ")
			'loComandoSeleccionar.AppendLine("	 FULL JOIN #tempPresupuestoProveedor2 ON #tempPresupuestoProveedor2.Cod_Art2 in (Select Cod_Art1 from #tempPresupuestoProveedor1) or #tempPresupuestoProveedor1.Cod_Art1 in (Select Cod_Art2 from #tempPresupuestoProveedor2)")
			'loComandoSeleccionar.AppendLine("	 FULL JOIN #tempPresupuestoProveedor3 ON #tempPresupuestoProveedor3.Cod_Art3 in (Select Cod_Art1 from #tempPresupuestoProveedor1) or #tempPresupuestoProveedor3.Cod_Art3 in (Select Cod_Art2 from #tempPresupuestoProveedor2)")
			
		 '   loComandoSeleccionar.AppendLine("	Select #temArticulos2.Cod_Art, #temArticulos2.Nom_Art, * From 	#tempFinal ")
		 '   loComandoSeleccionar.AppendLine("	JOIN #temArticulos2 On #temArticulos2.Row = #tempFinal.Row ")

			loComandoSeleccionar.AppendLine("	select  ")
			loComandoSeleccionar.AppendLine("		#temArticulos2.Cod_Art, ")
			loComandoSeleccionar.AppendLine("		#temArticulos2.Nom_Art, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Can_Art11 from #tempPresupuestoProveedor1 Where Cod_Art = #temArticulos2.Cod_Art),0) As Can_Art11, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Precio11 from #tempPresupuestoProveedor1 Where Cod_Art = #temArticulos2.Cod_Art),0) As Precio11, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Por_Des1 from #tempPresupuestoProveedor1 Where Cod_Art = #temArticulos2.Cod_Art),0) As Por_Des1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Mon_Net1 from #tempPresupuestoProveedor1 Where Cod_Art = #temArticulos2.Cod_Art),0) As Mon_Net1,  ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Can_Art12 from #tempPresupuestoProveedor2 Where Cod_Art = #temArticulos2.Cod_Art),0) As Can_Art12, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Precio12 from #tempPresupuestoProveedor2 Where Cod_Art = #temArticulos2.Cod_Art),0) As Precio12, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Por_Des2 from #tempPresupuestoProveedor2 Where Cod_Art = #temArticulos2.Cod_Art),0) As Por_Des2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Mon_Net2 from #tempPresupuestoProveedor2 Where Cod_Art = #temArticulos2.Cod_Art),0) As Mon_Net2,	 ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Can_Art13 from #tempPresupuestoProveedor3 Where Cod_Art = #temArticulos2.Cod_Art),0) As Can_Art13, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Precio13 from #tempPresupuestoProveedor3 Where Cod_Art = #temArticulos2.Cod_Art),0) As Precio13, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Por_Des3 from #tempPresupuestoProveedor3 Where Cod_Art = #temArticulos2.Cod_Art),0) As Por_Des3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Mon_Net3 from #tempPresupuestoProveedor3 Where Cod_Art = #temArticulos2.Cod_Art),0) As Mon_Net3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Cod_Pro1 from #tempPresupuestoProveedor1) ,'') As Cod_Pro1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Cod_Pro2 from #tempPresupuestoProveedor2) ,'') As Cod_Pro2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Cod_Pro3 from #tempPresupuestoProveedor3) ,'') As Cod_Pro3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Nom_Pro1 from #tempPresupuestoProveedor1) ,'') As Nom_Pro1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Nom_Pro2 from #tempPresupuestoProveedor2) ,'') As Nom_Pro2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Nom_Pro3 from #tempPresupuestoProveedor3) ,'') As Nom_Pro3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Dis_Imp1 from #tempPresupuestoProveedor1) ,'') As Dis_Imp1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Dis_Imp2 from #tempPresupuestoProveedor2) ,'') As Dis_Imp2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Dis_Imp3 from #tempPresupuestoProveedor3) ,'') As Dis_Imp3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Net_Enc1 from #tempPresupuestoProveedor1) ,0) As Mon_Net_Enc1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Net_Enc2 from #tempPresupuestoProveedor2) ,0) As Mon_Net_Enc2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Net_Enc3 from #tempPresupuestoProveedor3) ,0) As Mon_Net_Enc3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Bru_Enc1 from #tempPresupuestoProveedor1) ,0) As Mon_Bru_Enc1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Bru_Enc2 from #tempPresupuestoProveedor2) ,0) As Mon_Bru_Enc2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Bru_Enc3 from #tempPresupuestoProveedor3) ,0) As Mon_Bru_Enc3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Des_Enc1 from #tempPresupuestoProveedor1) ,0) As Mon_Des_Enc1, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Des_Enc2 from #tempPresupuestoProveedor2) ,0) As Mon_Des_Enc2, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Mon_Des_Enc3 from #tempPresupuestoProveedor3) ,0) As Mon_Des_Enc3, ")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Documento1 from #tempPresupuestoProveedor1) ,0) As Documento1,")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Documento2 from #tempPresupuestoProveedor2) ,0) As Documento2,")
			loComandoSeleccionar.AppendLine("		ISNULL((select Top 1 Documento3 from #tempPresupuestoProveedor3) ,0) As Documento3,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Dis_Imp1 from #tempPresupuestoProveedor1) ,0) As Dis_Imp1,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Dis_Imp2 from #tempPresupuestoProveedor2) ,0) As Dis_Imp2,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Dis_Imp3 from #tempPresupuestoProveedor3) ,0) As Dis_Imp3,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Comentario1 from #tempPresupuestoProveedor1) ,0) As Comentario1,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Comentario2 from #tempPresupuestoProveedor2) ,0) As Comentario2,")
			loComandoSeleccionar.AppendLine("	  	ISNULL((select Top 1 Comentario3 from #tempPresupuestoProveedor3) ,0) As Comentario3")

			loComandoSeleccionar.AppendLine("	from #temArticulos2  ")



            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

        '-------------------------------------------------------------------------------------------------------
        ' Verifica si el select (tabla nº0) trae registros
        '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
  
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAnalisis_compras", laDatosReporte)

			
		'-------------------------------------------------------------------------------------------------------
        ' Buscando la primera distribucion de impuestos
        '-------------------------------------------------------------------------------------------------------
			Dim lcXml1 As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto1 As String
            Dim loImpuestos1 As New System.Xml.XmlDocument()

            lcPorcentajesImpueto1 = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml1 = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Dis_Imp1")
						
                If String.IsNullOrEmpty(lcXml1.Trim()) Then
                    Continue For
                End If

                loImpuestos1.LoadXml(lcXml1)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos1.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ")"
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace("(,", "(")
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace(".", ",")
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto1.ToString
            
            
		'-------------------------------------------------------------------------------------------------------
        ' Buscando la segunda distribucion de impuestos
        '-------------------------------------------------------------------------------------------------------
            
            lcPorcentajesImpueto1 = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml1 = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Dis_Imp2")
						
                If String.IsNullOrEmpty(lcXml1.Trim()) Then
                    Continue For
                End If

                loImpuestos1.LoadXml(lcXml1)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos1.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ")"
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace("(,", "(")
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace(".", ",")
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto1.ToString
            
            
		'-------------------------------------------------------------------------------------------------------
        ' Buscando la tercera distribucion de impuestos
        '-------------------------------------------------------------------------------------------------------
            
            lcPorcentajesImpueto1 = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml1 = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Dis_Imp3")
						
                If String.IsNullOrEmpty(lcXml1.Trim()) Then
                    Continue For
                End If

                loImpuestos1.LoadXml(lcXml1)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos1.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto1 = lcPorcentajesImpueto1 & ")"
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace("(,", "(")
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace(".", ",")
            lcPorcentajesImpueto1 = lcPorcentajesImpueto1.Replace(".", ",")																					               
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto1.ToString


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfAnalisis_compras.ReportSource = loObjetoReporte

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
' RJG: 14/05/10: Codigo inicial (no terminado)												'
'-------------------------------------------------------------------------------------------'
