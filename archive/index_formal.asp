
<%@Language="VBScript" CodePage="65001"%>
<!--#include file="base.asp"-->
<%
'**********************************************
'下面两行是为了跳过微信的验证'
'response.write request("echostr")
'response.end
'**********************************************
dim mbx_wx=new mbx_weixin

mbx_wx.deal_xml
%>
