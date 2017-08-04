<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CGS_rTotales_HorasExtra.aspx.vb" Inherits="CGS_rTotales_HorasExtra" %>
<%@ Register Assembly="vis3Controles" Namespace="vis3Controles" TagPrefix="vis3Controles" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>Listado de Totales de Horas Extra (CGS)</title>
	<link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css" rel="stylesheet" type="text/css" />
</head>
<body>
	<form id="form1" runat="server">
		<div>
			<CR:CrystalReportViewer ID="crvCGS_rTotales_HorasExtra" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="False" EnableParameterPrompt="False" HasCrystalLogo="False" HasPrintButton="False" />
			<asp:ScriptManager ID="ScriptManager1" runat="server">
				<Scripts>
				</Scripts>
			</asp:ScriptManager>
			<asp:UpdatePanel ID="UpdatePanel1" runat="server">
				<ContentTemplate>
					<vis3Controles:wbcImpresoraReportes runat="server" ID="wbcImpresoraDeReportes" plMostrarBotonImprimir='True' />
					<vis3Controles:pnlVentanaModal ID="PnlVentanaModalPrincipal" runat="server" pcEstiloBotonCerrar="BotonCerrarVentanaModal" pcEstiloFondo="FondoVentanaModal" pcEstiloMarco="MarcoVentanaModal" pcTextoBotonCerrar="Cerrar" plMostrarBotonCerrar="false" poAlto="520px" poAncho="550px" Style="left: -24px; top: 61px" />
					<vis3Controles:pnlMensajeModal ID="PnlMensajeModal" runat="server" pcEstiloContenido="ContenidoMensajeModal" pcEstiloFondo="FondoVentanaModal" pcEstiloTitulo="TituloMensajeModal" pcEstiloVentana="MarcoMensajeModal" poAlto="400px" poAncho="750px" poArriba="20%" poIzquierda="30%" Style="left: -24px; top: 61px" />
					<vis3Controles:wbcAdministradorMensajeModal ID="WbcAdministradorMensajeModal" runat="server" Style="left: -24px; top: 61px" />
					<vis3Controles:wbcAdministradorVentanaModal ID="WbcAdministradorVentanaModal" runat="server" Style="left: -24px; top: 61px" />
				</ContentTemplate>
			</asp:UpdatePanel>
		</div>
	</form>
</body>
</html>
