<%@Language="VBScript" CodePage="65001"%>
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
        end if
        
Else
'strsend=text(fromusername,tousername,Content)
strsend=""
end if
response.write strsend
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
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/batchblacklist?access_token="&token
	dim str:str="{"
		str=str&"	""opened_list"":["
		str=str&"	"""&fromusername&"""]"
		str=str&"}"	
	dim http, strMsg
	set http=server.createobject("MSXML2.SERVERXMLHTTP.3.0")
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
			if http.readystate=4 and http.status=200 then 
			'if err.number = 0 then
			  strMsg="You are set to blacklist successfully!"
			  'strMsg=http.responseBody
			Else
			  strMsg="You are not set to blacklist!"
			End if	
		'strMsg="Enter"
		set_blacklist="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & strMsg & "]]></Content>" &_
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
