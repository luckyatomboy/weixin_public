<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	'option explicit
	session.codepage=65001
	response.charset="utf-8"
	server.scripttimeout=999999
	dim xmlhttp:xmlhttp="Msxml2.ServerXMLHTTP"
	dim xmldomString:xmldomString="MSXML2.DOMDocument"
%>
<!--#include file="constants.asp"-->
<!--#include file="class/js.asp"-->
<!--#include file="class/utility.asp"-->
<!--#include file="class/function.asp"-->