

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
dim Nickname     

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

		  'strsend=bp(FromUserName,ToUserName,news_beef_future, eventkey)
		  strsend=send_news(FromUserName,ToUserName,"猪肉报盘","","",news_beef_future)
		else
		  nickname=get_nickname(FromUserName)
		 strsend=post_template(FromUserName,ToUserName,nickname)		 
		end if
      end if

Elseif msgtype="text" then
'strsend=text(fromusername,tousername,Content)
	if Content=send_template_code then 
		strsend=get_userlist(FromUserName,ToUserName)
	else
		strsend=send_text(fromusername,tousername,Content)
	end if

end if

response.write strsend

set xml_dom=Nothing

'****************************************


%>
