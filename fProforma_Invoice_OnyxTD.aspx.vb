'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fProforma_Invoice_OnyxTD"
'-------------------------------------------------------------------------------------------'
Partial Class fProforma_Invoice_OnyxTD
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cotizaciones.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Generico, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Ent, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Rec1                       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For                        AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Notas, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Cotizaciones.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Cod_Uni2='') THEN Renglones_Cotizaciones.Can_Art1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Cotizaciones.Can_Art2 END) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Cod_Uni2='') THEN Renglones_Cotizaciones.Cod_Uni")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Cotizaciones.Cod_Uni2 END) AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Cod_Uni2='') THEN Renglones_Cotizaciones.Precio1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Cotizaciones.Precio1*Renglones_Cotizaciones.Can_Uni2 END) AS Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Imp1 As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           RTRIM(LTRIM(Articulos.Garantia)) As Garantia,")
            'loComandoSeleccionar.AppendLine("           Articulos.Peso, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Cod_Uni2='') THEN (Renglones_Cotizaciones.Can_Art1 * Articulos.Peso) ")
            loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Cotizaciones.Can_Art2 * Articulos.Peso) END) AS Peso, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Volumen, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Cotizaciones.Cod_Uni2='') THEN (Renglones_Cotizaciones.Can_Art1 * Articulos.Volumen) ")
            loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Cotizaciones.Can_Art2 * Articulos.Volumen) END) AS Volumen, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Ubi, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Pai, ")
            loComandoSeleccionar.AppendLine("           Paises.Nom_Pai,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Est, ")
            loComandoSeleccionar.AppendLine("           Estados.Nom_Est,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Ciu, ")
            loComandoSeleccionar.AppendLine("           Ciudades.Nom_Ciu,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Contacto, ")
            loComandoSeleccionar.AppendLine("           Clientes.Correo  ")
            loComandoSeleccionar.AppendLine(" FROM      Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Paises, ")
            loComandoSeleccionar.AppendLine("           Estados, ")
            loComandoSeleccionar.AppendLine("           Ciudades ")
            loComandoSeleccionar.AppendLine(" WHERE     Cotizaciones.Documento      =   Renglones_Cotizaciones.Documento ")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Tra    =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Paises.Cod_Pai = Clientes.Cod_Pai ")
            loComandoSeleccionar.AppendLine("           AND Estados.Cod_Est = Clientes.Cod_Est ")
            loComandoSeleccionar.AppendLine("           AND Ciudades.Cod_Ciu = Clientes.Cod_Ciu ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art   =   Renglones_Cotizaciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProforma_Invoice_OnyxTD", laDatosReporte)

            loObjetoReporte.SetParameterValue("Leyenda_Cotizaciones_Tipo1", goOpciones.mObtener("LEYCOTVEN1", "M"))
            loObjetoReporte.SetParameterValue("Leyenda_Cotizaciones_Tipo2", goOpciones.mObtener("LEYCOTVEN2", "M"))
            loObjetoReporte.SetParameterValue("Leyenda_Cotizaciones_Tipo3", goOpciones.mObtener("LEYCOTVEN3", "M"))

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfProforma_Invoice_OnyxTD.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal1.mMostrarMensajeModal("Error", _
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
' MAT: 09/03/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 14/03/11: Corrección de las unidades del reporte según requerimientos
'-------------------------------------------------------------------------------------------'
' MAT: 05/09/11: Creación de los parámetros para las leyendas en el formato
'-------------------------------------------------------------------------------------------'
' JFP: 09/10/12: Adecuacion a OnyxTD
'-------------------------------------------------------------------------------------------'
