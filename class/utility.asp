<%

'调用接口取得access_token'
function token_data()
	dim url:url="https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_appid&"&secret="&wx_secret
	dim json,b,c
dim http, strMsg, result
set http=server.createobject("MSXML2.SERVERXMLHTTP")
	http.open "GET",url,false
	http.send()		
	set json=toobject(http.responseText)

	b=json.access_token
	c=json.expires_in
	set_cache "get_token",b
	set_cache "expires_in",c
	set_cache "token_date",now()
	token_data=b

	set json=nothing
end function

'从buffer取得access token'
function get_token()
	if load_cache("get_token")="" then
		get_token=token_data
	else
		If datediff("s",load_cache("token_date"),now())>load_cache("expires_in") then
			get_token=token_data
		else
			get_token=load_cache("get_token")
		end if
	end if
end function

function set_cache(byval t0,byval t1)
		application.lock()
		application(prefix&t0)=t1
		application.unlock()
end function
	
function load_cache(byval t0)
  load_cache=application(prefix&t0)
end function	

function send_news(byval fromuser, byval touser, byval title, byval description, byval picurl, byval url)
send_news="<xml>" &_
"<ToUserName><![CDATA["&fromuser&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&touser&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[news]]></MsgType>" &_
"<ArticleCount>1</ArticleCount>" &_
"<Articles><item><Title><![CDATA["&title&"]]></Title>" &_
"<Description><![CDATA["&description&"]]></Description>" &_
"<PicUrl><!CDATA["&picurl&"]]></PicUrl>" &_
"<Url><![CDATA[" & url & "]]></Url></item></Articles>" &_
"</xml> "
end function

function send_text(byval fromuser, byval touser, byval content)
send_text="<xml>" &_
"<ToUserName><![CDATA["&fromuser&"]]></ToUserName>" &_
"<FromUserName><![CDATA["&touser&"]]></FromUserName>" &_
"<CreateTime>"&now&"</CreateTime>" &_
"<MsgType><![CDATA[text]]></MsgType>" &_
"<Content><![CDATA[" & content & "]]></Content>" &_
"</xml>"
end function

%>