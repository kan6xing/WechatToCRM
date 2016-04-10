var context = Xrm.Page.context;
var serverUrl = context.getServerUrl();
var getOrgUniqueName = context.getOrgUniqueName();

function CalculateTicketNumber() {
    //alert("opened");
    var new_number = Xrm.Page.getAttribute("new_number").getValue();
    //alert(new_number);
    var new_startcode = Xrm.Page.getAttribute("new_startcode").getValue();

    if (new_startcode != null)
        if (new_startcode.length != 14) {
            alert("票券号码必须为14位！");
            //alert(new_startcode.length);
        }

    if (new_number != null && new_startcode != null) {
        //alert(new_startcode);
        var Head = new_startcode.substr(0, 3);
        var Tail = new_startcode.substr(12);
        var StartNumber = new_startcode.substr(3, 9);
        var EndNumber = "00000000" + String(Number(StartNumber) + new_number - 1);

        Xrm.Page.getAttribute("new_startnumber").setValue(Tail + StartNumber);   //给计算查重字段赋值
        Xrm.Page.getAttribute("new_endnumber").setValue(Tail + EndNumber.substr(EndNumber.length - 9));       //给计算查重字段赋值

        EndNumber = Head + EndNumber.substr(EndNumber.length - 9) + Tail;
        //alert(EndNumber);
        Xrm.Page.getAttribute("new_endcode").setValue(EndNumber);

    }
}

function onIsCouponChange() {
    var new_iscoupon = Xrm.Page.getAttribute("new_iscoupon").getValue();
    if (new_iscoupon) {
        Xrm.Page.getAttribute("new_pricebase").setValue(1);

        //更新与产品相关的信息
        var order = Xrm.Page.getAttribute("new_salesorderid").getValue();
        var orderId = order[0].id;
        var Helper = new RESTHelper();
        LoadProductInf_Gift(orderId, Helper);
        Xrm.Page.getAttribute("new_closeprice").setValue(1);
    }
    else {
        var order = Xrm.Page.getAttribute("new_salesorderid").getValue();

        var orderId = order[0].id;
        var Helper = new RESTHelper();

        //取订单上的影城
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SalesOrderSet(guid'" +
          orderId + "')";
        var data = Helper.Read(queryurl);
        if (data != null) {
            //加载订单产品中的内容
            LoadProductInf(data.new_OrderType.Value, orderId, Helper);
        }
    }
}

function ini() {
    Xrm.Page.ui.controls.get("new_endcode").setDisabled(true);
    Xrm.Page.ui.controls.get("new_cinema").setDisabled(true);
    Xrm.Page.ui.controls.get("new_pricebase").setDisabled(true);
    Xrm.Page.ui.controls.get("new_closeprice").setDisabled(true);
    Xrm.Page.ui.controls.get("new_issuedate").setDisabled(true);
    Xrm.Page.ui.controls.get("new_expirydate").setDisabled(true);

    //创建时执行
    var order = Xrm.Page.getAttribute("new_salesorderid").getValue();
    if (order == null) {
        alert("影城不能为空！");
        return;
    }

    var orderId = order[0].id;
    if (Xrm.Page.ui.getFormType() == 1) {
        var Helper = new RESTHelper();

        //取订单上的影城
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SalesOrderSet(guid'" +
          orderId + "')";
        var data = Helper.Read(queryurl);
        if (data != null) {
            var cinemaLookup = {};
            cinemaLookup.id = data.new_Cinema.Id;
            cinemaLookup.entityType = 'new_cinema';
            cinemaLookup.name = data.new_Cinema.Name;
            var cinemaLookupValue = [];
            cinemaLookupValue[0] = cinemaLookup;

            if (cinemaLookupValue != null) {
                Xrm.Page.getAttribute("new_cinema").setValue(cinemaLookupValue);
            }

            Xrm.Page.getAttribute("new_name").setValue(data.Name);

            //加载订单产品中的内容
            LoadProductInf(data.new_OrderType.Value, orderId, Helper);
        }


    }
}

function onsave(ExecutionObj) {

    var new_startcode = Xrm.Page.getAttribute("new_startcode").getValue();

    if (Xrm.Page.getAttribute("new_cinema").getValue() == null) {
        alert("影城不能为空，请修改正确再保存！");
        ExecutionObj.getEventArgs().preventDefault();
    }
    else if (Xrm.Page.getAttribute("new_pricebase").getValue() == null) {
        alert("基础价不能为空，请修改正确再保存！");
        ExecutionObj.getEventArgs().preventDefault();
    }
    else if (new_startcode.length != 14) {
        alert("票券号码必须为14位，请修改正确再保存！");
        ExecutionObj.getEventArgs().preventDefault();
    }
    else if (!UpdatePOSPrice()) {
        alert("卖品抵用金额必须大于0，请修改正确再保存！");
        ExecutionObj.getEventArgs().preventDefault();
    }
    else if (!DateValid()) {
        alert("终止日期必须大于发行日期，请修改正确再保存！");
        ExecutionObj.getEventArgs().preventDefault();
    }
    else if (Xrm.Page.data.entity.getIsDirty()) {
        Xrm.Page.ui.controls.get("new_endcode").setDisabled(false);
        Xrm.Page.ui.controls.get("new_cinema").setDisabled(false);
        Xrm.Page.ui.controls.get("new_pricebase").setDisabled(false);
        Xrm.Page.ui.controls.get("new_closeprice").setDisabled(false);
        Xrm.Page.ui.controls.get("new_issuedate").setDisabled(false);
        Xrm.Page.ui.controls.get("new_expirydate").setDisabled(false);
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

function EnableAdd_B() {
    alert("gjhgkjhgk");
}
function EnableAdd() {
    return true;
}

//更新卖品抵用金额
function UpdatePOSPrice() {
    var type = Xrm.Page.getAttribute("new_useplace").getText();

    if (type == "GA") {
        Xrm.Page.getAttribute("new_posprice").setValue(0);
    }
    else {
        if (Number(Xrm.Page.getAttribute("new_posprice").getValue()) <= 0) {
            return false;
        }
    }

    return true;
}

//日期合法性判断
function DateValid() {
    var date1 = Xrm.Page.getAttribute("new_issuedate").getValue();
    var date2 = Xrm.Page.getAttribute("new_expirydate").getValue();

    if (date1 != null && date2 != null) {
        if (date2 >= date1) {
            return true;
        }
        else {
            return false;
        }
    }
    return false;
}

//加载订单产品中的有效期、基础价等
function LoadProductInf(orderType, orderId, Helper) {
    var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SalesOrderDetailSet?$filter=SalesOrderId/Id eq (guid'" + orderId + "')";
    var data = Helper.Read(queryurl);
    if (data != null) {
        for (i = 0; i < data.results.length; i++) {
            var productid = data.results[i].ProductId.Id;
            queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/ProductSet(guid'" +
                productid + "')";
            var product = Helper.Read(queryurl);
            if (product != null) {
                if (orderType == 100000000 || orderType == 100000001) {//A或B订单
                    if (product.new_ProductType.Value == 100000000 || product.new_ProductType.Value == 100000004) {
                        Xrm.Page.getAttribute("new_closeprice").setValue(Number(data.results[i].new_ClosePrice.Value));
                        Xrm.Page.getAttribute("new_pricebase").setValue(Number(data.results[i].new_ClosePrice.Value));  //new_BasePrice.Value

                        if (data.results[i].new_StartDate == null) return;
                        var d = new Date(Number(data.results[i].new_StartDate.substr(6, 13)));
                        Xrm.Page.getAttribute("new_issuedate").setValue(d);
                        d = new Date(Number(data.results[i].new_EndDate.substr(6, 13)));
                        Xrm.Page.getAttribute("new_expirydate").setValue(d);
                        return;
                    }
                }
                if (orderType == 100000002) {//C订单
                    if (product.new_ProductType.Value == 100000012) {
                        //Xrm.Page.getAttribute("new_pricebase").setValue(Number(data.results[i].new_BasePrice.Value));
                        var d = new Date(Number(data.results[i].new_StartDate.substr(6, 13)));
                        Xrm.Page.getAttribute("new_issuedate").setValue(d);
                        d = new Date(Number(data.results[i].new_EndDate.substr(6, 13)));
                        Xrm.Page.getAttribute("new_expirydate").setValue(d);
                        return;
                    }
                }
            }
        }
    }
}

//加载订单产品中赠券的有效期等
function LoadProductInf_Gift(orderId, Helper) {
    var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SalesOrderDetailSet?$filter=SalesOrderId/Id eq (guid'" + orderId + "')";
    var data = Helper.Read(queryurl);
    if (data != null) {
        for (i = 0; i < data.results.length; i++) {
            var productid = data.results[i].ProductId.Id;
            queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/ProductSet(guid'" +
                productid + "')";
            var product = Helper.Read(queryurl);
            if (product != null) {
                if (product.new_ProductType.Value == 100000011) {
                    Xrm.Page.getAttribute("new_closeprice").setValue(Number(data.results[i].new_ClosePrice.Value));

                    if (data.results[i].new_StartDate == null) return;

                    var d = new Date(Number(data.results[i].new_StartDate.substr(6, 13)));
                    Xrm.Page.getAttribute("new_issuedate").setValue(d);
                    d = new Date(Number(data.results[i].new_EndDate.substr(6, 13)));
                    Xrm.Page.getAttribute("new_expirydate").setValue(d);
                    return;
                }
            }
        }
    }
    alert("订单产品中无赠券，请先输入订单的赠券产品！");
}