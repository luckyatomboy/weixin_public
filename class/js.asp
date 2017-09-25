
<script language="JScript" runat="Server">

function ToObject(json) {
    var o;
    eval("o=" + json );
    return o;
    //eval("var o="+json);return o;
}

function getItemProperty(obj,Num,Name){
return obj[Num][Name];
}

function getItem(obj,Num){
return obj[Num];
}

</script>