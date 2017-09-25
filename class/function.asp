<%

function get_userlist(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim next_openid:next_openid=""
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/get?access_token="&access_token&"&next_openid="
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send()
		result=http.responseText
	dim nickname:nickname=get_nickname(fromusername)
		set http=nothing
		set json=toobject(result)
		for i=0 to json.data.openid.length-1
			userid=getitem(json.data.openid,i)
			'strmsg=get_nickname(fromusername,tousername,userid)&"/"&strmsg
			strmsg=post_template(userid,tousername,nickname)
		next
end function

function get_resourcelist(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token="&access_token
	dim str:str="" 
		str=str&"{"
		str=str&"""type"":""image"","
		str=str&"""offset"":0,"
		str=str&"""count"":2"
		str=str&"}"
	dim http, strMsg, result, itemtext
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		set json=toobject(result)
		for i=0 to json.item.length-1
			itemtext=getItemProperty(json.item,0,"name")
		next

		get_resourcelist=send_text(FromUserName, ToUserName, itemtext)
end function

function post_template(fromusername,tousername,byuser)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/message/template/send?access_token="&access_token
	dim str:str="" 
		str=str&"{"
		str=str&"""touser"":"
		str=str&""""&fromusername&""","
		str=str&"""template_id"":"""&template_id&""","
		str=str&"""url"":"""","		
		str=str&"""data"":{"
		str=str&"""first"":{""value"":""尊敬的用户，新的报盘已生成，请点击菜单查看\n\r(by"&byuser&")"",""color"":""#173177""},"
		str=str&"""keyword1"":{""value"":""一骥报盘""},"
		str=str&"""keyword2"":{""value"":""牛猪鸡期货现货报盘""}"
		'str=str&"""remark"":{""value"":""请点击菜单查看各品种报价""}"
		str=str&"}}"
	dim http, strMsg, result
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		post_template=""
end function

function get_nickname(userid)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid="&userid
	dim http, strMsg, result
	set http=server.createobject(xmlhttp)
		http.open "GET",url,false
		http.send()
		result=http.responsetext
'******Must use 'SET' to define JSON object********
		set json=ToObject(result)
		nickname=json.nickname
		'get_nickname=userid&":"&nickname
		get_nickname=nickname
end function

%>