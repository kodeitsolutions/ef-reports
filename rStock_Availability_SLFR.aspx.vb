'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices


'-------------------------------------------------------------------------------------------'
' Inicio de clase "rStock_Availability_SLFR"
'-------------------------------------------------------------------------------------------'
Partial Class rStock_Availability_SLFR
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcUsuario As String = goServicios.mObtenerCampoFormatoSQL(goUsuario.pcCodigo)
            loComandoSeleccionar.AppendLine("SELECT		COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-ALM'       AND Cod_Reg = " & lcUsuario & "), '') AS Almacenes,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-DPTO'      AND Cod_Reg = " & lcUsuario & "), '') AS Departamentos,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-VENDED'    AND Cod_Reg = " & lcUsuario & "), '') AS Vendedores,")
            loComandoSeleccionar.AppendLine("			COALESCE((SELECT	Val_Mem FROM Campos_propiedades WHERE Cod_Pro = 'USU-MARCAS'    AND Cod_Reg = " & lcUsuario & "), '') AS Marcas")

            loComandoSeleccionar.AppendLine("")

            Dim loPermisos As DataTable
            loPermisos = (New goDatos()).mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "Campos_Propiedades").Tables(0)
            loComandoSeleccionar.Length = 0

            Dim lcAlmacenesUsuario As String = CStr(loPermisos.Rows(0).Item("Almacenes")).Trim()
            Dim lcDepartamentosUsuario As String = CStr(loPermisos.Rows(0).Item("Departamentos")).Trim()
            Dim lcVendedoresUsuario As String = CStr(loPermisos.Rows(0).Item("Vendedores")).Trim()
            Dim lcMarcasUsuario As String = CStr(loPermisos.Rows(0).Item("Marcas")).Trim()

            lcAlmacenesUsuario = goServicios.mObtenerListaFormatoSQL(lcAlmacenesUsuario)
            lcDepartamentosUsuario = goServicios.mObtenerListaFormatoSQL(lcDepartamentosUsuario)
            lcVendedoresUsuario = goServicios.mObtenerListaFormatoSQL(lcVendedoresUsuario)
            lcMarcasUsuario = goServicios.mObtenerListaFormatoSQL(lcMarcasUsuario)


            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art,")
            loComandoSeleccionar.AppendLine("			SUM(")
            loComandoSeleccionar.AppendLine("				CASE")
            loComandoSeleccionar.AppendLine("					WHEN (DATEDIFF(day, GETDATE(), Ordenes_Compras.Fec_Fin ) <= 30) THEN Renglones_oCompras.can_pen1")
            loComandoSeleccionar.AppendLine("					ELSE 0")
            loComandoSeleccionar.AppendLine("				END")
            loComandoSeleccionar.AppendLine("			) AS Tran_Less,")
            loComandoSeleccionar.AppendLine("			SUM(")
            loComandoSeleccionar.AppendLine("				CASE")
            loComandoSeleccionar.AppendLine("					WHEN DATEDIFF(day, GETDATE(), Ordenes_Compras.Fec_Fin) > 30 THEN Renglones_oCompras.can_pen1")
            loComandoSeleccionar.AppendLine("					ELSE 0")
            loComandoSeleccionar.AppendLine("				END")
            loComandoSeleccionar.AppendLine("			) AS Tran_More")
            loComandoSeleccionar.AppendLine(" INTO #temporal")
            loComandoSeleccionar.AppendLine(" FROM Articulos")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_oCompras ON Articulos.Cod_Art = Renglones_oCompras.Cod_Art")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Ordenes_Compras ON Renglones_oCompras.Documento = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Ordenes_Compras.Status NOT IN ('Anulado','Procesado')")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art           Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("      	    And Articulos.Cod_Ubi between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	    And Articulos.Exi_Act1 >= " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_oCompras.Cod_Alm IN (" & lcAlmacenesUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep IN (" & lcDepartamentosUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar IN (" & lcMarcasUsuario & ")")
            loComandoSeleccionar.AppendLine(" GROUP BY Articulos.Cod_Art")

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Web,")
            loComandoSeleccionar.AppendLine("			Articulos.Promocion,")
            loComandoSeleccionar.AppendLine("			Articulos.Status,")
            loComandoSeleccionar.AppendLine("			Articulos.Tipo,")
            loComandoSeleccionar.AppendLine("			Articulos.Clase,")
            loComandoSeleccionar.AppendLine("			Articulos.Upc, ")
            loComandoSeleccionar.AppendLine("			Articulos.Abc,")
            loComandoSeleccionar.AppendLine("			Articulos.Modelo,")
            loComandoSeleccionar.AppendLine("			Articulos.Talla,")
            loComandoSeleccionar.AppendLine("			Articulos.Item, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Ubi,")
            loComandoSeleccionar.AppendLine("			Articulos.Ubicacion,")
            loComandoSeleccionar.AppendLine("			Articulos.Informacion,")
            loComandoSeleccionar.AppendLine("			Articulos.Ancho,")
            loComandoSeleccionar.AppendLine("			Articulos.Alto,")
            loComandoSeleccionar.AppendLine("			Articulos.Fondo,")
            loComandoSeleccionar.AppendLine("			Articulos.Peso,")
            loComandoSeleccionar.AppendLine("			Articulos.Volumen,")
            loComandoSeleccionar.AppendLine("			Articulos.Garantia,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Dep,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Sec,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Cla,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Tip,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni3, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni4, ")
            loComandoSeleccionar.AppendLine("			Articulos.Tip_uni,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Articulos.Color,")
            loComandoSeleccionar.AppendLine("			Articulos.Precio1,")
            loComandoSeleccionar.AppendLine("			Articulos.Precio2,")
            loComandoSeleccionar.AppendLine("			Articulos.Precio3,")
            loComandoSeleccionar.AppendLine("			Articulos.Precio4,")
            loComandoSeleccionar.AppendLine("			Articulos.Precio5,")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Imp,")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Imp,")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Ped1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Ped2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Cot1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Cot2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Pro1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Pro2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Pro1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Por2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Des1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Des2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Dis1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Dis2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Comentario,")
            loComandoSeleccionar.AppendLine("			Articulos.notas,")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Pro1, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Pro2, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Ult1, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Ult2, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Ant1, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Ant2, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Cli1, ")
            loComandoSeleccionar.AppendLine("			Articulos.CoS_Cli2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Pre,")
            loComandoSeleccionar.AppendLine("			Articulos.Foto,")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Gan1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Gan2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Gan3, ")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Gan4, ")
            loComandoSeleccionar.AppendLine("			Articulos.Por_Gan5, ")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Gan1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Gan2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Gan3, ")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Gan4, ")
            loComandoSeleccionar.AppendLine("			Articulos.Mon_Gan5, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Suc,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_imp,")
            loComandoSeleccionar.AppendLine("			Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine("			ISNULL(Unidades_Articulos.Can_Uni,Articulos.Can_Uni) AS Can_Uni,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Articulos.Exi_Act1 - Articulos.Exi_Ped1) BETWEEN 1 AND 10 THEN '-'")
            loComandoSeleccionar.AppendLine("				WHEN (Articulos.Exi_Act1 - Articulos.Exi_Ped1) BETWEEN 10 AND 50 THEN '10+'")
            loComandoSeleccionar.AppendLine("				WHEN (Articulos.Exi_Act1 - Articulos.Exi_Ped1) BETWEEN 50 AND 100 THEN '50+'")
            loComandoSeleccionar.AppendLine("				WHEN (Articulos.Exi_Act1 - Articulos.Exi_Ped1) BETWEEN 100 AND 500 THEN '100+'")
            loComandoSeleccionar.AppendLine("				WHEN (Articulos.Exi_Act1 - Articulos.Exi_Ped1) > 500 THEN '500+'")
            loComandoSeleccionar.AppendLine("				ELSE ' '")
            loComandoSeleccionar.AppendLine("			END AS Stock,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_Less BETWEEN 1 AND 10 THEN '-'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_Less BETWEEN 10 AND 50 THEN '10+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_Less BETWEEN 50 AND 100 THEN '50+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_Less BETWEEN 100 AND 500 THEN '100+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_Less > 500 THEN '500+'")
            loComandoSeleccionar.AppendLine("				ELSE ' '")
            loComandoSeleccionar.AppendLine("			END AS Tran_Less,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_More BETWEEN 1 AND 10 THEN '-'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_More BETWEEN 10 AND 50 THEN '10+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_More BETWEEN 50 AND 100 THEN '50+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_More BETWEEN 100 AND 500 THEN '100+'")
            loComandoSeleccionar.AppendLine("				WHEN #temporal.Tran_More > 500 THEN '500+'")
            loComandoSeleccionar.AppendLine("				ELSE ' '")
            loComandoSeleccionar.AppendLine("			END AS Tran_More")
            loComandoSeleccionar.AppendLine(" FROM Articulos")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Unidades_Articulos ON Articulos.Cod_Art = Unidades_Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Departamentos.Cod_Dep = Articulos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN #temporal ON Articulos.Cod_Art = #temporal.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Art           Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("      	    And Articulos.Cod_Ubi between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	    And Articulos.Exi_Act1 >= " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep IN (" & lcDepartamentosUsuario & ")")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar IN (" & lcMarcasUsuario & ")")
            If lcParametro8Desde = "Si" Then
                loComandoSeleccionar.AppendLine(" 		    And Cast(Articulos.Foto As VARCHAR) <> ''")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            laDatosReporte.Tables(0).Columns.Add("Foto2", GetType(String))
            laDatosReporte.Tables(0).Columns.Add("FotoImagen", GetType(Byte()))

            Dim lcXml As String = "<foto></foto>"
            Dim lcFoto As String = ""
            Dim lnNumeroImagenes As Integer = 0
            Dim loFotos As New System.Xml.XmlDocument()

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("foto")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loFotos.LoadXml(lcXml)
                lcFoto = "*"
                lnNumeroImagenes = 0

                'En cada renglón lee el contenido de cada imagen
                For Each loFoto As System.Xml.XmlNode In loFotos.SelectNodes("fotos/foto")
                    lcFoto = lcFoto & ", " & loFoto.SelectSingleNode("nombre").InnerText
                    lnNumeroImagenes = lnNumeroImagenes + 1
                Next loFoto

                lcFoto = lcFoto.Replace("*,", "")
                laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Foto2") = lnNumeroImagenes.ToString & lcFoto.ToString

            Next lnNumeroFila

            Me.mCargarFoto(laDatosReporte.Tables(0))

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rStock_Availability_SLFR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrStock_Availability_SLFR.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & " StackTrace: " & loExcepcion.StackTrace, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "600px", _
                           "500px")

        End Try

    End Sub

    Protected Sub mCargarFoto(ByRef loTabla As DataTable)

        ' Si la tabla no tiene registros
        If loTabla.Rows.Count <= 0 Then Return

        ' Se redimensiona la imagen 
        Dim loImage As Bitmap = Me.mRedimensionarImagen(MapPath(Me.pcLogoEmpresa), 50, 50)
        ' se carga en memoria
        Dim loMemory As MemoryStream = New MemoryStream()
        loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
        ' se guarda la imagen en un arreglo de byte
        Dim loImageByteEmpresa As Byte() = loMemory.GetBuffer()

        ' Recorriendo los registros de la tabla
        For j As Integer = 0 To (loTabla.Rows.Count - 1)

            ' Si el registro tiene imagen asociada
            If loTabla.Rows(j).Item("Foto2").ToString <> "" Then

                ' se extrae los nombres de archivo de imagen del registro
                Dim LcNombreImagen As String = loTabla.Rows(j).Item("Foto2").ToString.Substring(1)
                Dim LnNumeroImagenes As Integer = CInt(loTabla.Rows(j).Item("Foto2").ToString.Substring(0, 1))

                Dim lcMatrizNombres As New ArrayList()
                lcMatrizNombres.AddRange(Split(LcNombreImagen, ","))

                ' Si existe archivos de imagen asociado
                If LnNumeroImagenes > 0 Then

                    ' Recorriendo la lista de archivos de imagenes
                    For i As Integer = 0 To (lcMatrizNombres.Count - 1)

                        ' se eliminan los espacios en blanco
                        lcMatrizNombres(i) = lcMatrizNombres(i).ToString.ToUpper.Trim

                        ' Si existe el archivo de imagen
                        If IO.File.Exists(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString)) Then

                            ' Se redimensiona la imagen
                            loImage = Me.mRedimensionarImagen(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString), 50, 50)
                            ' se carga en memoria
                            loMemory = New MemoryStream()
                            loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
                            ' se guarda la imagen en un arreglo de byte
                            Dim loImageByte As Byte() = loMemory.GetBuffer()
                            ' se escribe en la tabla de registro
                            loTabla.Rows(j).Item("FotoImagen") = loImageByte

                        Else

                            '' Se redimensiona la imagen
                            'loImage = Me.mRedimensionarImagen(MapPath("../../FrameWork/Imagenes/SinImagen.png"), 50, 50)
                            '' se carga en memoria
                            'loMemory = New MemoryStream()
                            'loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
                            '' se guarda la imagen en un arreglo de byte
                            'Dim loImageByte As Byte() = loMemory.GetBuffer()
                            '' se escribe en la tabla de registro
                            'loTabla.Rows(j).Item("FotoImagen") = loImageByte
                            loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa

                        End If

                    Next

                Else

                    loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa

                End If
            Else

                ' se escribe en la tabla de registro
                loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa

            End If
        Next

    End Sub

    Protected Function mRedimensionarImagen(ByVal lcFilename As String, ByVal lnWidth As Integer, ByVal lnHeight As Integer) As Bitmap

        ' Se lee el archivo de la imagen
        Dim loArchivoImagen As IO.FileStream = New IO.FileStream(lcFilename, IO.FileMode.Open, IO.FileAccess.Read)
        ' Se carga la imagen
        Dim loBMP As Bitmap = New Bitmap(loArchivoImagen)
        ' Variable donde se guardar la imagen redimensionada
        Dim bmpOut As Bitmap = New Bitmap(lnWidth, lnHeight)
        Try

            Dim lnRatio As Decimal
            Dim lnNewWidth As Integer = 0
            Dim lnNewHeight As Integer = 0

            ' Si el tamaño de la imagen es menor a la que se quiere redimensionar
            If (loBMP.Width < lnWidth And loBMP.Height < lnHeight) Then
                ' se retorna la imagen original
                Return loBMP
            End If

            ' Si el ancho de la imagen original es mayo que la altura de la imagen original
            If (loBMP.Width > loBMP.Height) Then
                ' se calcula la relacion de anchura para redimensionar
                lnRatio = lnWidth / loBMP.Width
                ' ancho de la nueva imagen
                lnNewWidth = lnWidth
                ' se calcula la altura de la nueva imagen
                Dim lnTemp As Decimal = loBMP.Height * lnRatio
                lnNewHeight = lnTemp
            Else
                ' se calcula la relacion de altura para redimensionar
                lnRatio = lnHeight / loBMP.Height
                ' altura de la nueva imagen
                lnNewHeight = lnHeight
                ' se calcula la anchura de la nueva imagen
                Dim lnTemp As Decimal = loBMP.Width * lnRatio
                lnNewWidth = lnTemp
            End If

            ' se crea la imagen nueva para redimensionar
            bmpOut = New Bitmap(lnNewWidth, lnNewHeight, loBMP.PixelFormat)
            ' se carga la manipulacion de la imagen
            Dim g As Graphics = Graphics.FromImage(bmpOut)
            ' se estable el modo de interpolacion de la imagen para redimensionar
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            ' se carga el tamaño al que se redimensionara
            g.FillRectangle(Brushes.White, 0, 0, lnNewWidth, lnNewHeight)
            ' se dibuja la imagen redimensionandola
            g.DrawImage(loBMP, 0, 0, lnNewWidth, lnNewHeight)

            loBMP.Dispose()
        Catch
            ' si ocurre un error, retorna la imagen original
            Return loBMP

        End Try
        ' retorna la imagen redimensionada
        Return bmpOut

    End Function

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
' DLC: 12/04/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 02/07/2010: Código para redimensionar las imagenes de los artículos, con el fin
'                   de minimizar el peso del archivo.
'-------------------------------------------------------------------------------------------' 
' DLC: 16/07/2010: Se ajusto la consulta de la base de datos, agregando un nuevo filtro
'                   "Quitar Stocks menores a: "
'-------------------------------------------------------------------------------------------' 
' DLC: 23/09/2010: - Se ajusto la selección por rango de fecha para determinar si esta 
'                   en el rango de menos de 30 días o mayores a 30 dias.
'                   - Se ajusto la visualización, correspondiendo a los datos enviados.
'                   (Si las columnas no aparece nada es porque no hay en stock ni en ordenes de compra)
'                   - La selección de los articulos en las ordenes de compra se basa en 
'                   el estatus de la orden, si esta anulada o procesada se descarta
'-------------------------------------------------------------------------------------------' 
' MAT: 11/01/2011: Correción y mantenimiento del Reporte
'-------------------------------------------------------------------------------------------'
' JFP: 30/08/2012: Adecuacion para el nuevo reporte filtrado de la Fuerza de Venta
'-------------------------------------------------------------------------------------------'

