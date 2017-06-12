<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TRF_rEGanancias_Perdidas.aspx.vb" Inherits="TRF_rEGanancias_Perdidas" %>

<%@ Register Assembly="vis2Controles" Namespace="vis2Controles" TagPrefix="vis2Controles" %>
<%@ Register Assembly="vis1Controles" Namespace="vis1Controles" TagPrefix="vis1Controles" %>
<%@ Register Assembly="vis3Controles" Namespace="vis3Controles" TagPrefix="vis3Controles" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>


    

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Estado de Ganancias y P�rdidas (TRF)</title>
    <link href="~/Framework/cssEstilosFramework.css" rel="stylesheet" type="text/css" />
    <link href="~/Administrativo/cssEstilosAdministrativo.css" rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <CR:CrystalReportViewer ID="crvTRF_rEGanancias_Perdidas" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="False"
            EnableParameterPrompt="False" HasCrystalLogo="False" />
       <vis3Controles:pnlVentanaModal ID="PnlVentanaModalPrincipal" runat="server" pcEstiloBotonCerrar="BotonCerrarVentanaModal"
            pcEstiloFondo="FondoVentanaModal" pcEstiloMarco="MarcoVentanaModal" pcTextoBotonCerrar="Cerrar"
            plMostrarBotonCerrar="false" poAlto="520px" poAncho="550px" Style="left: -16px;
            top: 50px" />
        <vis3Controles:pnlMensajeModal ID="PnlMensajeModalPrincipal" runat="server" pcEstiloContenido="ContenidoMensajeModal"
            pcEstiloFondo="FondoVentanaModal" pcEstiloTitulo="TituloMensajeModal" pcEstiloVentana="MarcoMensajeModal"
            poAlto="400px" poAncho="750px" poArriba="20%" poIzquierda="30%" />
        <vis3Controles:wbcAdministradorMensajeModal ID="WbcAdministradorMensajeModal" runat="server" />
        <vis3Controles:wbcAdministradorVentanaModal ID="WbcAdministradorVentanaModal" runat="server" />
    </div>
    </form>
</body>
</html>
