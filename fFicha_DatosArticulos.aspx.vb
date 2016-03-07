'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosArticulos"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosArticulos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Art,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine(" 	CASE")
			loComandoSeleccionar.AppendLine(" 		WHEN Articulos.Status = 'A' THEN 'Acttivos'")
			loComandoSeleccionar.AppendLine(" 		WHEN Articulos.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 		WHEN Articulos.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 	END AS Status,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Modelo,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Abc,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Tip,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Dep,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Cla,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Nacional,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Usa_ser,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Usa_lot,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Upc,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Sec,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Pro,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Mar,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Precio1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Precio2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Precio3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Precio4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Precio5,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pre1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pre2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pre3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pre4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pre5,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Gan1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Gan2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Gan3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Gan4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Gan5,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Gan1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Gan2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Gan3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Gan4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Gan5,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Gan_Min,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Ali,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Ara,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Imp,")
			'loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Tip,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Pro1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Pro2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Ult1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Ult2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Ant1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Ant2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Cli1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Cli2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Pro,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Ult,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Ant,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Cli,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Tipo,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cos_Pre,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Act1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Act2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Ped1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Ped2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Cot1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Cot2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Pro1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Pro2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Por1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Por2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Des1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Des2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Dis1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Dis2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi5,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Exi6,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Max,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Min,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Pto,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Exi_Otr,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Can_Uni,")
			loComandoSeleccionar.AppendLine(" 	CASE ")
			loComandoSeleccionar.AppendLine(" 		WHEN Articulos.Tip_Uni = 'U' THEN 'Única'")
			loComandoSeleccionar.AppendLine(" 		WHEN Articulos.Tip_Uni = 'M' THEN 'Multiple'")
			loComandoSeleccionar.AppendLine(" 	END AS Tip_Uni,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Tie_Fab,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Uni1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Uni2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Imp_Lic,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Comentario,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Peso,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Volumen,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Color,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Garantia,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Gra_Lic,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cap_Lic,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Ubi,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Pre_Nac,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Ubicacion,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Talla,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Generico,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Max_Ven,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Min_Ven,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Item,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Tie_Min,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Tie_Max,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Tie_Pto,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Vid_Uti,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Ancho,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Alto,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Fondo,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Notas,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Informacion,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Cod_con,		")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Otr1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Otr2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Otr3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Otr4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Mon_Otr5,	 ")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Otr1,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Otr2,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Otr3,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Otr4,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Por_Otr5,	  ")
			loComandoSeleccionar.AppendLine(" 	Articulos.Min_Com,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Max_Com,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Min_Pro,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Max_Pro,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Min_Aju,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Max_Aju,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Min_Tra,")
			loComandoSeleccionar.AppendLine(" 	Articulos.Max_Tra,")
			
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Cod_Uni,")
			loComandoSeleccionar.AppendLine(" Unidades.nom_uni,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Tipo AS Tipo_Unidad,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Can_uni,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Precio1 AS Precio1_Unidad,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Precio2 AS Precio2_Unidad,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Precio3 AS Precio3_Unidad,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Precio4 AS Precio4_Unidad,")
			loComandoSeleccionar.AppendLine(" Unidades_Articulos.Precio5 AS Precio5_Unidad,")
			loComandoSeleccionar.AppendLine(" Tipos_Articulos.Nom_tip AS Nombre_tipo_Art,")
			loComandoSeleccionar.AppendLine(" Clases_Articulos.Nom_Cla AS Nombre_clase_Art,")
			loComandoSeleccionar.AppendLine(" Departamentos.Nom_Dep,")
			loComandoSeleccionar.AppendLine(" Secciones.Nom_Sec,")
			loComandoSeleccionar.AppendLine(" Conceptos.Nom_Con,")
			loComandoSeleccionar.AppendLine(" Marcas.Nom_Mar,")
			loComandoSeleccionar.AppendLine(" Proveedores.Nom_Pro,")
			loComandoSeleccionar.AppendLine(" Unidad_Principal.Nom_Uni As Nombre_Unidad_Principal,")
			loComandoSeleccionar.AppendLine(" Impuestos.Nom_Imp")
			
			
			loComandoSeleccionar.AppendLine(" FROM Articulos")
			loComandoSeleccionar.AppendLine(" JOIN Tipos_Articulos ON Tipos_Articulos.Cod_Tip = Articulos.Cod_Tip")
			loComandoSeleccionar.AppendLine(" JOIN Clases_Articulos ON Clases_Articulos.Cod_Cla = Articulos.Cod_Cla")
			loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Departamentos.Cod_Dep = Articulos.Cod_Dep")
			loComandoSeleccionar.AppendLine(" JOIN Secciones ON Secciones.Cod_Sec = Articulos.Cod_Sec AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Unidades_Articulos ON Unidades_Articulos.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Unidades ON Unidades.Cod_Uni = Unidades_Articulos.Cod_Uni")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Conceptos ON Conceptos.Cod_Con = Articulos.Cod_Con")
			loComandoSeleccionar.AppendLine(" JOIN Marcas ON Marcas.Cod_Mar = Articulos.Cod_Mar")
			loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Articulos.Cod_Pro")
			loComandoSeleccionar.AppendLine(" JOIN Unidades AS Unidad_Principal ON Unidad_Principal.Cod_Uni = Articulos.Cod_Uni1")
			loComandoSeleccionar.AppendLine(" JOIN Impuestos ON Impuestos.Cod_Imp = Articulos.Cod_Imp")
			
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
			
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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosArticulos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosArticulos.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 22/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 11/05/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'