<%@Language="VBScript" CodePage="936"%>

<%
'**********************************************
'ע������
'ASP�ļ���Ҫ��UTF-8�ĸ�ʽ����,��������.
'�������д�����Ϊ��ͨ��΢�Žӿ���֤�ġ�
'response.write request("echostr")
'response.end
'**********************************************
dim signature        '΢�ż���ǩ��
dim timestamp        'ʱ���
dim nonce                '�����
'dim echostr                '����ַ���
dim Token
dim signaturetmp
token="Mybx2017Yjfood"'���ں�̨��д�� token

signature = Request("signature")
nonce = Request("nonce")
timestamp = Request("timestamp")
'**********************************************
dim ToUserName        '������΢�ź�
dim FromUserName'���ͷ��ʺţ�һ��OpenID��
dim CreateTime        '��Ϣ����ʱ�䣨���ͣ�
dim MsgType                'text
dim Content                '�ı���Ϣ����

set xml_dom = Server.CreateObject("MSXML2.DOMDocument")'�˴���������ʵ�ʷ����������д
xml_dom.load request
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text '������΢���˺�
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text '������΢���˺š������ǵĹ���ƽ̨�˺š�
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
if MsgType="text" then
Content=xml_dom.getelementsbytagname("Content").item(0).text
end if
'dim  mingling
'mingling=replace(content,chr(13),"")
'mingling=trim(replace(mingling,chr(10),""))
if (MsgType="event") then
        strEventType=xml_dom.getelementsbytagname("Event").item(0).text '΢���¼�
        if strEventType="subscribe" then '��ʾ����΢�Ź���ƽ̨
        	strMsg="Welcome subscribe MYBX!"
        ElseIf strEventType="unsubscribe" Then'ȡ����
        	strMsg="�ź�ȡ����ע��ζţׯ��"
        end if
        strsend=gz(FromUserName,ToUserName,strMsg)
Else
'strsend=text(fromusername,tousername,Content)
strsend=""
end if
response.write strsend
set xml_dom=Nothing
'*************���´���ֻ��Ϊ�˵�������***********
'filepath=server.mappath(".")&"\wx.txt"
'Set fso = Server.CreateObject("Scripting.FileSystemObject")
'set fopen=fso.OpenTextFile(filepath, 8 ,true)
'fopen.writeline(strsend)
'set fso=nothing
'set fopen=Nothing
'****************���Խ���************************

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
