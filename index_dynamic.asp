

<!--#include file="base.asp"-->
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
dim EventKey         

set xml_dom = Server.CreateObject(xmldomString)
xml_dom.load request
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text 
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text 
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text

if MsgType="text" then
	Content=xml_dom.getelementsbytagname("Content").item(0).text
elseif MsgType="event" then
	EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
end if

if MsgType="event" then
        strEventType=xml_dom.getelementsbytagname("Event").item(0).text 
     if strEventType="CLICK" then
		if EventKey="beef" then
		  'strsend=get_resourcelist(FromUserName,ToUserName)
		  strsend=get_userlist(FromUserName,ToUserName)
		elseif EventKey="pork" then
		  strMsg="http://mp.weixin.qq.com/s?__biz=MzIyNjgzODc4Ng==&mid=100000015&idx=1&sn=6f1bc275683ebd847d2e5ceb1593a233&chksm=686b11625f1c9874ce54ecf8e2eb3ef8b6b0d78d6361715587340475d71ffea6f7a99ae7479e#rd"
		  'strsend=pic_bp(FromUserName,ToUserName,strMsg, eventkey)
		  strsend=bp(FromUserName,ToUserName,strMsg, eventkey)
		
		 'strsend=get_nickname(FromUserName,ToUserName)
  		  'strsend=get_resourcelist(FromUserName,ToUserName)
		else
		 strsend=dyn_post_template(FromUserName,ToUserName)		 
		end if
      end if

Elseif msgtype="text" then
'strsend=text(fromusername,tousername,Content)
	if Content="mybxbp2017" then 
		strsend=get_userlist(FromUserName,ToUserName)
	else
		strsend=text(fromusername,tousername,Content)
	end if

end if

response.write strsend

set xml_dom=Nothing

'****************************************
function dyn_post_template(fromusername,tousername)
	dim access_token
	access_token=get_token()
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
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		dyn_post_template=""
end function


function get_userlist(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim next_openid:next_openid=""
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/get?access_token="&access_token&"&next_openid="
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send()
		result=http.responseText
		set http=nothing
		set json=toobject(result)
		for i=0 to json.data.openid.length-1
			userid=getitem(json.data.openid,i)
			'strmsg=get_nickname(fromusername,tousername,userid)&"/"&strmsg
			strmsg=post_template(userid,tousername)
		next
end function

function get_resourcelist(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token="&access_token
	dim str:str="" 'str="{""opened_list"":[""oAqdVwDJNEo9-HBbuxe4ZyMYptqo"",""oAqdVwDU44MFiwOqxi-mnuKXLIhQ""]}"
		str=str&"{"
		str=str&"""type"":""image"","
		str=str&"""offset"":0,"
		str=str&"""count"":2"
		str=str&"}"
	dim http, strMsg, result, itemtext
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		set json=toobject(result)
		for i=0 to json.item.length-1
			itemtext=getItemProperty(json.item,0,"name")
		next

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
	dim access_token
	access_token=get_token()
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
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		post_template=""
end function

function get_nickname(fromusername,tousername,userid)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid="&userid
	dim http, strMsg, result
	set http=server.createobject("MSXML2.SERVERXMLHTTP")
		http.open "GET",url,false
		http.send()
		result=http.responsetext
'******Must use 'SET' to define JSON object********
		set json=ToObject(result)
		nickname=json.nickname
		get_nickname=userid&":"&nickname
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
