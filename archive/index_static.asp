
<%@Language="VBScript" CodePage="65001"%>
<!--#include file="class/js.asp"-->
<%
'**********************************************
'下面两行是为了跳过微信的验证'
'response.write request("echostr")
'response.end
'**********************************************
dim signature        
dim timestamp       
dim nonce               
'dim echostr                '
dim Token
dim signaturetmp
token="Mybx2017Yjfood"

signature = Request("signature")
nonce = Request("nonce")
timestamp = Request("timestamp")
'**********************************************
dim ToUserName       
dim FromUserName
dim CreateTime      
dim MsgType               
dim Content            

set xml_dom = Server.CreateObject("MSXML2.DOMDocument")
xml_dom.load request
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text 
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text 
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
if MsgType="text" then
Content=xml_dom.getelementsbytagname("Content").item(0).text
end if
'dim  mingling
'mingling=replace(content,chr(13),"")
'mingling=trim(replace(mingling,chr(10),""))
if (MsgType="event") then
        strEventType=xml_dom.getelementsbytagname("Event").item(0).text 
     if strEventType="subscribe" then '如果是新用户关注了公众号，先把他设为黑名单  	
        	strsend=set_blacklist(fromusername,tousername)
        ElseIf strEventType="unsubscribe" Then
        	strMsg="Sorry for that!"
        	strsend=gz(FromUserName,ToUserName,strMsg)
     elseif strEventType="CLICK" then
		if EventKey="beef" then
		  'strMsg="http://shmbx.iego.cn/pic/beef.jpg"
		  'strsend=set_blacklist(fromusername,tousername)
		  strsend=get_resourcelist(FromUserName,ToUserName)
		elseif EventKey="pork" then
		  'strMsg="http://www.yj-food.com/weixin/jellyfish.jpg"
		  strMsg="http://mp.weixin.qq.com/s?__biz=MzIyNjgzODc4Ng==&mid=100000015&idx=1&sn=6f1bc275683ebd847d2e5ceb1593a233&chksm=686b11625f1c9874ce54ecf8e2eb3ef8b6b0d78d6361715587340475d71ffea6f7a99ae7479e#rd"
		  'strsend=pic_bp(FromUserName,ToUserName,strMsg, eventkey)
		  strsend=bp(FromUserName,ToUserName,strMsg, eventkey)
		
		 'strsend=get_nickname(FromUserName,ToUserName)
  		  'strsend=get_resourcelist(FromUserName,ToUserName)
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

'************************
'filepath=server.mappath(".")&"\wx.txt"
'Set fso = Server.CreateObject("Scripting.FileSystemObject")
'set fopen=fso.OpenTextFile(filepath, 8 ,true)
'fopen.writeline(strsend)
'set fso=nothing
'set fopen=Nothing
'****************************************
function set_blacklist(fromusername,tousername)
	dim access_token:access_token="6yYvMY-ytFmUIOcokye00QjkU80CEBq49c3J_uhl2Nngi-HTSwilwffeZ6wrjPDJROf-b1n6e7B4F1Wzua4_HapdH4zXErMxZKUDy6f2SkK9jkH2fCnoU_UjsUjfCdrTFFTaAJAVVJ"
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
		"<Content><![CDATA[黑名单" & fromusername & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function get_resourcelist(fromusername,tousername)
	dim access_token:access_token="girQuAcUU8CRRyjRQKYN2ZQgQAzdxttguAefhO_npdyoUvsahfPWHRumipOZ35rm-hY0Ydbg0shou0y78x2-9osCsNAYHIWyYmr_Rx8pMPzpQCcfGo7aktX_BxtU7QaREURbABAQGD"
	dim url:url="https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""type"":""image"","
		str=str&"""offset"":0,"
		str=str&"""count"":1"
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
		set json=toobject(result)
		for i=0 to json.item.length
			'mediaItem=ToObject(json.item(i))
			'itemtext=json.item(i)
			count=i
		next
		'mediaItem=json.item(1)
		'strMsg="Enter"
		get_resourcelist="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & result & "]]></Content>" &_
		"<FuncFlag>0</FuncFlag>" &_
		"</xml> "
end function

function post_template(fromusername,tousername)
	dim access_token:access_token="girQuAcUU8CRRyjRQKYN2ZQgQAzdxttguAefhO_npdyoUvsahfPWHRumipOZ35rm-hY0Ydbg0shou0y78x2-9osCsNAYHIWyYmr_Rx8pMPzpQCcfGo7aktX_BxtU7QaREURbABAQGD"
	dim url:url="https://api.weixin.qq.com/cgi-bin/message/template/send?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""touser"":"
		str=str&""""&fromusername&""","
		str=str&"""template_id"":""75zwrz_nM1M3Stb25P2FSuJs_w0NKHtFuK_8ewPGLa0"","
		'str=str&"""url"":""http://db.auto.sina.com.cn/photo/c99499-99-2.html#129958895"","
		str=str&"""url"":""http://mp.weixin.qq.com/s?__biz=MzIyNjgzODc4Ng==&mid=100000015&idx=1&sn=6f1bc275683ebd847d2e5ceb1593a233&chksm=686b11625f1c9874ce54ecf8e2eb3ef8b6b0d78d6361715587340475d71ffea6f7a99ae7479e#rd"","
		str=str&"""data"":{"
		str=str&"""first"":{""value"":""报盘""},"
		str=str&"""keyword1"":{""value"":""名称""},"
		str=str&"""keyword2"":{""value"":""简介""},"
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
	dim access_token:access_token="h9Cvtj71fMfkLzN9MdMbojLdWQa8iNZhKccT_EpqSou_9x1kBD6w2lCaLykZFElhTC5yZZIYzHeQnRv8iVGdemu32KhbB3vXmrn81K7aLpyCFLG_tJClDxEdFQ2XmP6OYYUdAGAQQM"
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
	dim access_token:access_token="6yYvMY-ytFmUIOcokye00QjkU80CEBq49c3J_uhl2Nngi-HTSwilwffeZ6wrjPDJROf-b1n6e7B4F1Wzua4_HapdH4zXErMxZKUDy6f2SkK9jkH2fCnoU_UjsUjfCdrTFFTaAJAVVJ"
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

function bp(fromusername,tousername,strmsg, eventkey)
bp="<xml>" &_
"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[news]]></MsgType>" &_
"<ArticleCount>1</ArticleCount>" &_
"<Articles><item><Title><![CDATA["&eventkey&"]]></Title>" &_
"<Description><![CDATA[beefaaa]]></Description>" &_
"<PicUrl><!CDATA[]]></PicUrl>" &_
"<Url><![CDATA[" & strmsg & "]]></Url></item></Articles>" &_
"</xml> "
end function

function pic_bp(fromusername,tousername,strmsg, eventkey)
pic_bp="<xml>" &_
"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[image]]></MsgType>" &_
"<Image>" &_
"<MediaId><![CDATA[rsxbqc9uiM0n2q9pgLsl0zUKxKdN3kpvmwIdRnyf5Fo]]></MediaId>" &_
"</Image>" &_
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
