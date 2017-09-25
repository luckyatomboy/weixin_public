
<%@Language="VBScript" CodePage="65001"%>
<!--#include file="class/js.asp"-->
<%
'**********************************************
'注意事项
'ASP文件需要以UTF-8的格式保�?否则乱码.
'以下两行代码是为了通过微信接口验证的�?
'response.write request("echostr")
'response.end
'**********************************************
dim signature        '微信加密签名
dim timestamp        '时间�?
dim nonce                '随机�?
'dim echostr                '随机字符�?
dim Token
dim signaturetmp
token="Mybx2017Yjfood"'您在后台添写�?token

signature = Request("signature")
nonce = Request("nonce")
timestamp = Request("timestamp")
'**********************************************
dim ToUserName        '开发者微信号
dim FromUserName'发送方帐号（一个OpenID�?
dim CreateTime        '消息创建时间（整型）
dim MsgType                'text
dim Content                '文本消息内容

set xml_dom = Server.CreateObject("MSXML2.DOMDocument")'此处根据您的实际服务器情况改�?
xml_dom.load request
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text '发送者微信账�?
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text '接收者微信账号。即我们的公众平台账号�?
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
if MsgType="text" then
Content=xml_dom.getelementsbytagname("Content").item(0).text
end if
'dim  mingling
'mingling=replace(content,chr(13),"")
'mingling=trim(replace(mingling,chr(10),""))
if (MsgType="event") then
        strEventType=xml_dom.getelementsbytagname("Event").item(0).text '微信事件
     if strEventType="subscribe" then '表示订阅微信公众平台      	
        	strsend=set_blacklist(fromusername,tousername)
        ElseIf strEventType="unsubscribe" Then'取消�?
        	strMsg="Sorry for that!"
        	strsend=gz(FromUserName,ToUserName,strMsg)
     elseif strEventType="CLICK" then
		if EventKey="beef" then
		  'strMsg="http://shmbx.iego.cn/pic/beef.jpg"
		  strsend=set_blacklist(fromusername,tousername)
		elseif EventKey="pork" then
		  'strMsg="http://shmbx.iego.cn/pic/others.jpg"
		
		 'strsend=get_nickname(FromUserName,ToUserName)
  		  strsend=get_resourcelist(FromUserName,ToUserName)
		else
		 strsend=post_template(FromUserName,ToUserName)		 
		end if
		'strsend=bp(FromUserName,ToUserName,strMsg)
	'strsend=set_blacklist(fromusername,tousername)
        	'strMsg=EventKey
        	'strsend=gz(FromUserName,ToUserName,strMsg)
        end if
        
Else
'strsend=text(fromusername,tousername,Content)
strsend=""
end if
if strsend<>"" then
response.write strsend
end if
set xml_dom=Nothing

'*************以下代码只是为了调试作用***********
'filepath=server.mappath(".")&"\wx.txt"
'Set fso = Server.CreateObject("Scripting.FileSystemObject")
'set fopen=fso.OpenTextFile(filepath, 8 ,true)
'fopen.writeline(strsend)
'set fso=nothing
'set fopen=Nothing
'****************调试结束************************
function set_blacklist(fromusername,tousername)
	dim access_token:access_token="m119n1qS6hxFLXOJWYQMYfRxj8izU5QrPBW7-viAohP5Ax7aKgiUaz6bRlpORpCPBGuGwy3w9vkhLFSBVhoB3MCywjH9BPJ37i-Hjykd9JgHKEZDrv12z3KG2cQvzPZ9PBWeAHALXS"
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/batchblacklist?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""opened_list"":["
		'str=str&""""&fromusername&"""]"
		str=str&"""ghghghghg""]"
		str=str&"}"
	dim http, strMsg, result
	'set http=server.createobject(xmlhttp)
	'set http=server.createobject("MSXML2.XMLHTTP")
	'set http=server.createobject("MSXML2.SERVERXMLHTTP.3.0")
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
			'if http.readystate=4 and http.status=200 then 
				'result=http.responseText
			'if err.number = 0 then
			  'strMsg="You are set to blacklist successfully!"
			'Else
				'result=http.responseText
			  'strMsg="You are not set to blacklist!"
			'End if	
		result=http.responseText
		'result="{""errcode"":0,""errmsg"":""ok""}"
		set http=nothing
		set json=ToObject(result)
		'strMsg="Enter"
		set_blacklist="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & fromusername & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function get_resourcelist(fromusername,tousername)
	dim access_token:access_token="m119n1qS6hxFLXOJWYQMYfRxj8izU5QrPBW7-viAohP5Ax7aKgiUaz6bRlpORpCPBGuGwy3w9vkhLFSBVhoB3MCywjH9BPJ37i-Hjykd9JgHKEZDrv12z3KG2cQvzPZ9PBWeAHALXS"
	dim url:url="https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""type"":""news"","
		'str=str&"""offset"":""7"","
		str=str&"""count"":""1"""
		str=str&"}"
	dim http, strMsg, result
	'set http=server.createobject(xmlhttp)
	'set http=server.createobject("MSXML2.XMLHTTP")
	'set http=server.createobject("MSXML2.SERVERXMLHTTP.3.0")
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
			'if http.readystate=4 and http.status=200 then 
				'result=http.responseText
			'if err.number = 0 then
			  'strMsg="You are set to blacklist successfully!"
			'Else
				'result=http.responseText
			  'strMsg="You are not set to blacklist!"
			'End if	
		result=http.responseText
		'result="{""errcode"":0,""errmsg"":""ok""}"
		set http=nothing
		set json=ToObject(result)
		for i=0 to json.item.size
			itemtext=json.item(i).media_id
		next
		'strMsg="Enter"
		get_resourcelist="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & itemtext & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function post_template(fromusername,tousername)
	dim access_token:access_token="m119n1qS6hxFLXOJWYQMYfRxj8izU5QrPBW7-viAohP5Ax7aKgiUaz6bRlpORpCPBGuGwy3w9vkhLFSBVhoB3MCywjH9BPJ37i-Hjykd9JgHKEZDrv12z3KG2cQvzPZ9PBWeAHALXS"
	dim url:url="https://api.weixin.qq.com/cgi-bin/message/template/send?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""touser"":"
		str=str&""""&fromusername&""","
		str=str&"""template_id"":""75zwrz_nM1M3Stb25P2FSuJs_w0NKHtFuK_8ewPGLa0"","
		str=str&"""url"":""https://mp.weixin.qq.com/s?__biz=MzIyNjgzODc4Ng==&tempkey=OTE5X2ViYk1naEVvYkpPazgvMXYxNzV3Q2tVSlRXT3pjNUE5TUU4NkJCcUFBQzdfeXBpUWJ1bUtQVTZJUGNHS0h1RGR0bVNrdUlDZ0t4Z1hFRDIwcTQteEZ5MTVLbGVaaVdZbUpwenFNbWdISVpxVGdpemxweGZiaVBGcjF6OE5NeGN4ZzI4UVYwc2lUQ0JVaENING00UXd0QnAwRVlPSDhVTFh3bGwyb0F%2Bfg%3D%3D&chksm=686b11755f1c9863ed3bd491916d17c72c9c11b51ffcd03a2551020037d8ecd983dc6dd12121#rd"","
		str=str&"""data"":{"
		str=str&"""first"":{""value"":""baopan""},"
		str=str&"""keyword1"":{""value"":""mingcheng""},"
		str=str&"""keyword2"":{""value"":""jianjie""},"
		str=str&"""remark"":{""value"":""remark""}"
		str=str&"}}"
	dim http, strMsg, result
	'set http=server.createobject(xmlhttp)
	'set http=server.createobject("MSXML2.XMLHTTP")
	'set http=server.createobject("MSXML2.SERVERXMLHTTP.3.0")
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
			'if http.readystate=4 and http.status=200 then 
				'result=http.responseText
			'if err.number = 0 then
			  'strMsg="You are set to blacklist successfully!"
			'Else
				'result=http.responseText
			  'strMsg="You are not set to blacklist!"
			'End if	
		result=http.responseText
		'result="{""errcode"":0,""errmsg"":""ok""}"
		set http=nothing
		post_template=""
end function

function get_blacklist(fromusername,tousername)
	dim access_token:access_token="m119n1qS6hxFLXOJWYQMYfRxj8izU5QrPBW7-viAohP5Ax7aKgiUaz6bRlpORpCPBGuGwy3w9vkhLFSBVhoB3MCywjH9BPJ37i-Hjykd9JgHKEZDrv12z3KG2cQvzPZ9PBWeAHALXS"
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/getblacklist?access_token="&access_token
	dim str:str="" 
		str=str&"{"
		str=str&"""begin_openid"":}"
	dim http, strMsg, result
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
			'if http.readystate=4 and http.status=200 then 
				'result=http.responseText
			'if err.number = 0 then
			  'strMsg="You are set to blacklist successfully!"
			'Else
				'result=http.responseText
			  'strMsg="You are not set to blacklist!"
			'End if	
		result=http.responsetext
		'result="({""errcode"":0,""errmsg"":""ok""})"
		'set http=nothing
'******Must use 'SET' to define JSON object********
		set json=ToObject(result)
		nickname=json.data.openid
		'strMsg="Enter"
		get_blacklist="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & nickname & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function get_nickname(fromusername,tousername)
	dim access_token:access_token="m119n1qS6hxFLXOJWYQMYfRxj8izU5QrPBW7-viAohP5Ax7aKgiUaz6bRlpORpCPBGuGwy3w9vkhLFSBVhoB3MCywjH9BPJ37i-Hjykd9JgHKEZDrv12z3KG2cQvzPZ9PBWeAHALXS"
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid=oAqdVwDU44MFiwOqxi-mnuKXLIhQ"
	dim http, strMsg, result
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "GET",url,false
		'http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send()
			'if http.readystate=4 and http.status=200 then 
				'result=http.responseText
			'if err.number = 0 then
			  'strMsg="You are set to blacklist successfully!"
			'Else
				'result=http.responseText
			  'strMsg="You are not set to blacklist!"
			'End if	
		result=http.responsetext
		'result="({""errcode"":0,""errmsg"":""ok""})"
		'set http=nothing
'******Must use 'SET' to define JSON object********
		set json=ToObject(result)
		nickname=json.nickname
		'strMsg="Enter"
		get_nickname="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & nickname & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function bp(fromusername,tousername,strmsg,eventkey)
bp="<xml>" &_
"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[news]]></MsgType>" &_
"<Content><![CDATA[]]></Content>" &_
"<ArticleCount>1</ArticleCount>" &_
"<Articles><item><Title><![CDATA["&eventkey&"]]></Title>" &_
"<Description>><![CDATA[]]></Description>" &_
"<PicUrl><!CDATA[" & strmsg & "]]></PicUrl>" &_
"<Url>><![CDATA[" & strmsg & "]]></Url></item></Articles>" &_
"<FuncFlag>0</FuncFlag>" &_
"</xml> "
end function

function gz(fromusername,tousername,strmsg)
gz="<xml>" &_
"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[text]]></MsgType>" &_
"<Content><![CDATA[" & strmsg & "]]></Content>" &_
"<FuncFlag>0</FuncFlag>" &_
"</xml> "
end function

function text(fromusername,tousername,returnstr)
text="<xml>" &_
"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[text]]></MsgType>" &_
"<Content><![CDATA[" & returnstr & "]]></Content>" &_
"<FuncFlag>0<FuncFlag>" &_
"</xml>"
end function
%>
