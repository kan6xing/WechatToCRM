var context = Xrm.Page.context;
var serverUrl = context.getServerUrl();
var getOrgUniqueName = context.getOrgUniqueName();

function ini() {
        //Xrm.Page.ui.controls.get("new_inuse").setDisabled(true);
}

function onsave() {
    //if (Xrm.Page.data.entity.getIsDirty()) {
        //Xrm.Page.ui.controls.get("new_inuse").setDisabled(false);
    //}
}

function Activate() {
    var bValid=false;
    var roles=context.getUserRoles();
    for (i = 0; i < roles.length; i++) {
        var id = roles[i];
        var Helper = new RESTHelper();

        //查找角色
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/RoleSet(guid'" +
          id + "')";
        var data = Helper.Read(queryurl);
        if (data != null) {
            //是否为目标角色
            if (data.Name == "GH_启用价目表") {
                bValid = true;
            }
        }
    }

    if (!bValid) {
        alert("您无权启用价目表");
    }
    else {
        var inUse = Xrm.Page.getAttribute("new_inuse");
        if (inUse == null) {
            alert("缺少启用状态字段，无法启用");
            return false;
        }
        else {
            inUse.setValue(true);
        }

        Xrm.Page.data.entity.save();
    }
}

function Deactivate() {
    var bValid = false;
    var roles = context.getUserRoles();
    for (i = 0; i < roles.length; i++) {
        var id = roles[i];
        var Helper = new RESTHelper();

        //查找角色
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/RoleSet(guid'" +
          id + "')";
        var data = Helper.Read(queryurl);
        if (data != null) {
            //是否为目标角色
            if (data.Name == "GH_启用价目表") {
                bValid = true;
            }
        }
    }

    if (!bValid) {
        alert("您无权停用价目表");
    }
    else {
        var inUse = Xrm.Page.getAttribute("new_inuse");
        if (inUse == null) {
            alert("缺少启用状态字段，无法启用");
            return false;
        }
        else {
            inUse.setValue(false);
        }

        Xrm.Page.data.entity.save();
    }
}

/* 
    作者：hyx   
    简介：REST中的CRUD操作辅助脚本。 
*/
function RESTHelper() { }

/* 
    方法简介：通过REST对Dynamics CRM 中的实体进行Create操作。 
    输入参数： 
    createurl:调用Dynamics CRM数据服务的URL字符串。例如："/GH2011/XRMServices/2011/OrganizationData.svc/ContactSet" 
    jsondata：需要进行Create操作的对象，必须进行json序列化。 
    输出参数： 
    true：Create成功。 
    false：Create失败。 
*/
RESTHelper.prototype.Create = function (createurl, jsondata) {
    var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
    xmlhttp.Open("POST", createurl, false);
    xmlhttp.SetRequestHeader("Content-Type", "application/json; charset=utf-8");
    xmlhttp.SetRequestHeader("Content-Length", jsondata.length);
    xmlhttp.SetRequestHeader("Accept", "application/json");
    xmlhttp.Send(jsondata);

    if (xmlhttp.readyState == 4) {
        if (xmlhttp.status == 201) {
            return true;

        }
        else {
            return false;
        }
    }//if  
}


/* 
 
   方法简介：通过REST对Dynamics CRM 中的实体进行Read操作。 
   输入参数： 
        parameter:调用Dynamics CRM数据服务的URL字符串。例如："/GH2011/XRMServices/2011/OrganizationData.svc/ContactSet(guid'{B75B220A-D2A4-48F4-8002-D8B564A866EA}')" 
   输出参数： 
        Object:获得了返回值 
        Null:查询失败。 
*/
RESTHelper.prototype.Read = function (queryurl) {
    var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
    xmlhttp.Open("GET", queryurl, false);
    xmlhttp.SetRequestHeader("Content-Type", "application/json; charset=utf-8");
    xmlhttp.SetRequestHeader("Content-Length", 0);
    xmlhttp.SetRequestHeader("Accept", "application/json");
    xmlhttp.Send(null);

    if (xmlhttp.readyState == 4) {
        if (xmlhttp.status == 200) {
            return window.JSON.parse(xmlhttp.responseText).d;

        }
        else {
            return null;
        }
    }
}



/* 
简介:通过REST方式更新实体。 
输入参数描述: 
    updateurl:"/GH2011/XRMServices/2011/OrganizationData.svc/OpportunitySet(guid'{DA83B96B-DBAF-4F0C-A75D-7203F2502087}')" 
    entity:    需要更新的对象，对象必须为JASON格式。 
输出参数: 
    更新成功返回:true 
    更新失败返回:false 
*/
RESTHelper.prototype.Update = function (updateurl, entity) {
    var uptXmlHttpReq = new XMLHttpRequest();
    uptXmlHttpReq.open("POST", updateurl, false);
    uptXmlHttpReq.setRequestHeader("Accept", "application/json");
    uptXmlHttpReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    uptXmlHttpReq.setRequestHeader("X-HTTP-Method", "MERGE");

    uptXmlHttpReq.send(entity);

    if (uptXmlHttpReq.readyState == 4) {
        if (uptXmlHttpReq.status == 204 || uptXmlHttpReq.status == 1223) {
            return true;
        }
        else {
            return false;
        }
    }

}


/* 
简介:通过REST方式删除实体。 
参数描述: 
   deleteurl:"/GH2011/XRMServices/2011/OrganizationData.svc/ContactSet(guid'{DA83B96B-DBAF-4F0C-A75D-7203F2502087}')" 
返回类型: 
   删除成功返回:true 
   删除失败返回:false 
*/
RESTHelper.prototype.Delete = function (deleteurl) {
    var uptXmlHttpReq = new XMLHttpRequest();
    uptXmlHttpReq.open("POST", deleteurl, false); //第三个参数表示是否已异步的方式发起请求  
    uptXmlHttpReq.setRequestHeader("Accept", "application/json");
    uptXmlHttpReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    uptXmlHttpReq.setRequestHeader("X-HTTP-Method", "DELETE");

    uptXmlHttpReq.send(null);

    if (uptXmlHttpReq.readyState == 4) {
        if (uptXmlHttpReq.status == 204 || uptXmlHttpReq.status == 1223) {
            return true;
        }
        else {
            return false;
        }
    }

}