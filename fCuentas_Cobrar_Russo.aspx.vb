'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCuentas_Cobrar_Russo"
'-------------------------------------------------------------------------------------------'
Partial Class fCuentas_Cobrar_Russo

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT	Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("  UPPER(Tipos_Documentos.Nom_Tip) AS Nom_Tip, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Factura, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Status, ")
            loComandoSeleccionar.AppendLine("  CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Ini, 103)                                 AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("  CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Fin, 103)                                 AS	Fec_Fin, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_Tra                                                          AS  Cod_Tra, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_For                                                          AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Cod_Mon                                                          AS  Cod_Mon, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Tasa                                                             AS  Tasa, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Mon_Bru                                                          AS  Mon_Bru,")
            loComandoSeleccionar.AppendLine("  (Cuentas_Cobrar.Mon_Imp1 + Cuentas_Cobrar.Mon_Imp2 + Cuentas_Cobrar.Mon_Imp3)   AS  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Mon_Net                                                          AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Mon_Sal                                                          AS  Saldo, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Por_Rec                                                          AS  Por_Rec, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Mon_Rec                                                          AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Por_Des                                                          AS  Por_Des, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Mon_Des                                                          AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("  (Cuentas_Cobrar.Mon_Otr1 + Cuentas_Cobrar.Mon_Otr2 + Cuentas_Cobrar.Mon_Otr3)   AS  Mon_Otr, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Comentario                                                       AS  Comentario, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("  Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("  Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("  Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("  Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("  Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("  SUBSTRING(Clientes.Telefonos,1,15)                                              AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("  Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("  Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("  Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("  SUBSTRING(RTRIM(Transportes.Cod_Tra) +' '+ Transportes.Nom_Tra,1,30) AS Nom_Tra, ")
            loComandoSeleccionar.AppendLine("  Zona_Vendedores.nom_zon as Zon_Ven, ")
            loComandoSeleccionar.AppendLine("  Zona_Clientes.nom_zon as Zon_Cli, ")
            loComandoSeleccionar.AppendLine("  Facturas.Control AS Num_Con_Fac, ")
            loComandoSeleccionar.AppendLine("  Facturas.Fec_Ini AS Fec_Ini_Fac, ")
            loComandoSeleccionar.AppendLine("  Facturas.Mon_net AS Mon_net_Fac, ")
            loComandoSeleccionar.AppendLine("  Ciudades.Nom_Ciu ")
            loComandoSeleccionar.AppendLine("  FROM      Cuentas_Cobrar as Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("  JOIN Clientes	  ON Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("  JOIN Vendedores	  ON Cuentas_Cobrar.Cod_Ven  =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("  JOIN Formas_Pagos ON Cuentas_Cobrar.Cod_For  =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("  JOIN Transportes  ON Cuentas_Cobrar.Cod_Tra  =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("  JOIN Zonas   as Zona_Vendedores     ON Vendedores.Cod_Zon      =   Zona_Vendedores.Cod_Zon  ")
            loComandoSeleccionar.AppendLine("  JOIN Zonas   as Zona_Clientes     ON Clientes.Cod_Zon      =   Zona_Clientes.Cod_Zon ")
            loComandoSeleccionar.AppendLine("  LEFT JOIN Facturas as Facturas ON Facturas.Documento = Cuentas_Cobrar.Doc_Ori ")
            loComandoSeleccionar.AppendLine("  JOIN Tipos_Documentos On Tipos_Documentos.Cod_Tip = Cuentas_Cobrar.Cod_Tip ")
            loComandoSeleccionar.AppendLine("  JOIN Ciudades ON Ciudades.Cod_Ciu = Clientes.Cod_Ciu ")
            loComandoSeleccionar.AppendLine("  WHERE          " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Cobrar_Russo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Cobrar_Russo.ReportSource = loObjetoReporte

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
