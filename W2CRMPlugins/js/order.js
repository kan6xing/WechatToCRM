var context = Xrm.Page.context;
var serverUrl = context.getServerUrl();
var getOrgUniqueName = context.getOrgUniqueName();

//订单实体初始事件
function order_onLoad() {
    /*
    在新建订单时，将“销售员”字段初始化为当前用户，将“乙方影城”字段初始化为销售员的影城将影城置为不可编辑状态
    并将价目表设置为影城的价目表，并将“影城”及“价目表”置为不可编辑状态；如销售员上无影城信息，则置“价目表”为可编辑状态，“影城”置为不可编辑状态；
    
    在更新订单时，将“影城”置为不可编辑状态；如果影城不为空，则置“价目表”为不可编辑状态，反之，则置“价目表”为可编辑状态
    */

    //调用设置权限方法(调用xhx的方法)
    Authority();
    //调用框架协议显示与隐藏
    changeProposalnumber();
    if (Xrm.Page.ui.getFormType() == 1) {

        var userId = context.getUserId();



        var Helper = new RESTHelper();
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + userId + "')";
        var cinemanameID = "";

        var data = Helper.Read(queryurl);
        if (data != null) {
            var userName = data.FullName;
            var cinemaname = data.new_CinemaName;
            if (userName != null) {
                var userLookup = {};
                userLookup.id = userId;
                userLookup.entityType = 'SystemUser';
                userLookup.name = userName;
                var userLookupValue = [];
                userLookupValue[0] = userLookup;

                if (userLookupValue != null) {
                    Xrm.Page.getAttribute("new_saler").setValue(userLookupValue);
                }

            }

            if (cinemaname != null && cinemaname.Name != null) {
                var cinemaLookup = {};
                cinemaLookup.id = cinemaname.Id;
                cinemanameID = cinemaname.Id;
                cinemaLookup.entityType = 'new_cinema';
                cinemaLookup.name = cinemaname.Name;
                var cinemaLookupValue = [];
                cinemaLookupValue[0] = cinemaLookup;

                if (cinemaLookupValue != null) {
                    Xrm.Page.getAttribute("new_cinema").setValue(cinemaLookupValue);
                }
            }
        }

        //查询影城的价格表,并给订单的价格表字段赋值
        if (Xrm.Page.getAttribute("new_cinema").getValue() != null) {

            queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/new_cinemaSet(guid'" + Xrm.Page.getAttribute("new_cinema").getValue()[0].id + "')";

            data = Helper.Read(queryurl);
            if (data != null && data.new_PriceList != null) {
                var priceLookup = {};
                priceLookup.id = data.new_PriceList.Id;
                priceLookup.entityType = 'pricelevel';
                priceLookup.name = data.new_PriceList.Name;
                var priceLookupValue = [];
                priceLookupValue[0] = priceLookup;

                //Add by Changjun.Yao at 2013-11-15
                queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/PriceLevelSet(guid'" + data.new_PriceList.Id + "')";

                data = Helper.Read(queryurl);

                if (data.new_inuse != null)
                    if (data.new_inuse == true) {

                        //End of Adding by Changjun.Yao at 2013-11-15

                        //给订单的价格表字段赋值
                        Xrm.Page.getAttribute("pricelevelid").setValue(priceLookupValue);
                    }

            }
                //“价目表”置为不可编辑状态
                Xrm.Page.getControl("pricelevelid").setDisabled(false);


            if (data != null) {

                //给开户行赋值  new_AccountOpeningBank  new_cinemaopeningbank
                //给开户行赋值  new_AccountOpeningBank  new_cinemaopeningbank
                if (data.new_AccountOpeningBank != null) {
                    Xrm.Page.getAttribute("new_cinemaopeningbank").setValue(data.new_AccountOpeningBank);
                }
                else {
                    Xrm.Page.getAttribute("new_cinemaopeningbank").setValue("");
                }

                //给联系电话赋值  new_Phone  new_cinemaphone
                if (data.new_Phone != null) {
                    Xrm.Page.getAttribute("new_cinemaphone").setValue(data.new_Phone);
                }
                else {
                    Xrm.Page.getAttribute("new_cinemaphone").setValue("");
                }

                //收款银行名称赋值 new_CollectingBankName  new_cinemabank
                if (data.new_CollectingBankName != null) {
                    Xrm.Page.getAttribute("new_cinemabank").setValue(data.new_CollectingBankName);
                }
                else {
                    Xrm.Page.getAttribute("new_cinemabank").setValue("");
                }

                //帐号赋值     new_Account   new_cinemaaccountnumber
                if (data.new_Account != null) {
                    Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue(data.new_Account);
                }
                else {
                    Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue("");
                }

            }



        }
        else {

            //如销售员上无影城信息，则置“价目表”为可编辑状态          
            Xrm.Page.getControl("pricelevelid").setDisabled(true);

            Xrm.Page.getAttribute("new_cinemaopeningbank").setValue("");
            Xrm.Page.getAttribute("new_cinemaphone").setValue("");
            Xrm.Page.getAttribute("new_cinemabank").setValue("");
            Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue("");
        }



    }
    else if (Xrm.Page.ui.getFormType() == 2) {
        //设置影城只读
        Xrm.Page.getControl("new_cinema").setDisabled(true);

        //如果影城不为空设置价目表只读。否则设置价目表可以修改。
        if (Xrm.Page.getAttribute("new_cinema").getValue() != null) {

            Xrm.Page.getControl("pricelevelid").setDisabled(true);
        }
        else {

            Xrm.Page.getControl("pricelevelid").setDisabled(false);
        }

    }

    //“错误提示”字段的显示，如果此字段有值则显示，否则不显示
    if (Xrm.Page.getAttribute("new_error").getValue() != null && Xrm.Page.getAttribute("new_error").getValue() != "") {
        Xrm.Page.getControl("new_error").setVisible(true);

    }


    //将影城置为不可编辑状态
    Xrm.Page.getControl("new_cinema").setDisabled(true);

}

//订单实体保存事件
function order_onSave1(ExecutionObj) {
    /*
	订单类型为B时，框架协议编码不能为空，订单附件不能为空
	*/
    if (!Xrm.Page.data.entity.getIsDirty()) {
        ExecutionObj.getEventArgs().preventDefault();
        return;
    }


    Xrm.Page.getControl("new_cinema").setDisabled(false);


    var orderType = Xrm.Page.getAttribute("new_ordertype").getText();


    //框架协议编码
    var proposalnumber = Xrm.Page.getAttribute("new_proposalnumber").getValue();
    if (proposalnumber == null && orderType == 'B订单') {
        alert("订单类型等于B订单时，要求必须填写订单协议，如果未填写订单协议，将不允许保存订单!");
        ExecutionObj.getEventArgs().preventDefault();
        return;
    }
    //应收日期
    var receivabledate = Xrm.Page.getAttribute("new_receivabledate").getValue();
    if (receivabledate == null && orderType == 'B订单') {
        alert("订单类型等于B订单时，要求必须填写应收日期，如果未填写应收日期，将不允许保存订单!");
        ExecutionObj.getEventArgs().preventDefault();
        return;
    }

    if (Xrm.Page.ui.getFormType() == 2 && orderType == 'B订单') {


        var orderId = Xrm.Page.data.entity.getId();
        //附件  http://172.16.32.219:5555/ghcrmtest/XRMServices/2011/OrganizationData.svc/AnnotationSet?$filter=ObjectId/Id eq (guid'"+orderId+"')
        ////http://172.16.32.219:5555/ghcrmtest/XRMServices/2011/OrganizationData.svc/AnnotationSet?$filter=ObjectId/Id eq (guid'fb247ec5-a4e3-e211-a4d0-00259033ee4d')

        var Helper = new RESTHelper();
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/AnnotationSet?$filter=ObjectId/Id eq (guid'" + orderId + "')";
        var returnValue = Helper.Read(queryurl);
        if (returnValue == null || returnValue.results.length == 0) {
            alert("订单类型等于B订单时，要求必须添加订单相关附件，如果未添加附件，将不允许保存订单!");
            ExecutionObj.getEventArgs().preventDefault();
            // return false;
        }

    }

}


//订单实体销售员onChange事件
function new_saler_onChange() {
    setDepartment();

}


//订单实体客户onChange事件
function customerid_onChange() {
    /*
	选择客户后带出客户的地址赋给“甲方地址”字段
	*/
    setAddress();

}

//设置影城函数
function setDepartment() {
    /*
    影城赋值方法：根据销售员，取出销售员的影城信息赋给“乙方影城”字段
    */

    var userId = Xrm.Page.getAttribute("new_saler").getValue();
    var cinemaLookup = {};
    //为空判断
    if (userId != null) {

        var Helper = new RESTHelper();
        var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + userId[0].id + "')";

        var data = Helper.Read(queryurl);

        if (data != null && data.new_CinemaName.Name != null) {
            var cinemaname = data.new_CinemaName;

            cinemaLookup.id = cinemaname.Id;
            cinemaLookup.entityType = 'new_cinema';
            cinemaLookup.name = cinemaname.Name;
            var cinemaLookupValue = [];
            cinemaLookupValue[0] = cinemaLookup;
        }

    }


    Xrm.Page.getAttribute("new_cinema").setValue(cinemaLookupValue);

    //查询影城的价格表,并给订单的价格表字段赋值
    if (Xrm.Page.getAttribute("new_cinema").getValue() != null) {

        queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/new_cinemaSet(guid'" + Xrm.Page.getAttribute("new_cinema").getValue()[0].id + "')";

        data = Helper.Read(queryurl);
        if (data != null && data.new_PriceList != null) {
            var priceLookup = {};
            priceLookup.id = data.new_PriceList.Id;
            priceLookup.entityType = 'pricelevel';
            priceLookup.name = data.new_PriceList.Name;
            var priceLookupValue = [];
            priceLookupValue[0] = priceLookup;

            //给订单的价格表字段赋值
            Xrm.Page.getAttribute("pricelevelid").setValue(priceLookupValue);


            //“价目表”置为不可编辑状态
            Xrm.Page.getControl("pricelevelid").setDisabled(false);
        }


        if (data != null) {

            //给开户行赋值  new_AccountOpeningBank  new_cinemaopeningbank
            //给开户行赋值  new_AccountOpeningBank  new_cinemaopeningbank
            if (data.new_AccountOpeningBank != null) {
                Xrm.Page.getAttribute("new_cinemaopeningbank").setValue(data.new_AccountOpeningBank);
            }
            else {
                Xrm.Page.getAttribute("new_cinemaopeningbank").setValue("");
            }

            //给联系电话赋值  new_Phone  new_cinemaphone
            if (data.new_Phone != null) {
                Xrm.Page.getAttribute("new_cinemaphone").setValue(data.new_Phone);
            }
            else {
                Xrm.Page.getAttribute("new_cinemaphone").setValue("");
            }

            //收款银行名称赋值 new_CollectingBankName  new_cinemabank
            if (data.new_CollectingBankName != null) {
                Xrm.Page.getAttribute("new_cinemabank").setValue(data.new_CollectingBankName);
            }
            else {
                Xrm.Page.getAttribute("new_cinemabank").setValue("");
            }

            //帐号赋值     new_Account   new_cinemaaccountnumber
            if (data.new_Account != null) {
                Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue(data.new_Account);
            }
            else {
                Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue("");
            }

        }



    }
    else {

        //如销售员上无影城信息，则置“价目表”为可编辑状态          
        Xrm.Page.getControl("pricelevelid").setDisabled(true);

        Xrm.Page.getAttribute("new_cinemaopeningbank").setValue("");
        Xrm.Page.getAttribute("new_cinemaphone").setValue("");
        Xrm.Page.getAttribute("new_cinemabank").setValue("");
        Xrm.Page.getAttribute("new_cinemaaccountnumber").setValue("");
    }




}

//设置地址函数
function setAddress() {
    /*
			
	地址赋值方法：根据客户值，找到客户的地址赋给“甲方地址”字段
	*/
    Xrm.Page.getAttribute("new_accountaddress").setValue("");

    var accountId = Xrm.Page.getAttribute("customerid").getValue();

    if (accountId == null) {
        return;
    }

    var Helper = new RESTHelper();
    var queryurl = "/" + getOrgUniqueName + "/XRMServices/2011/OrganizationData.svc/AccountSet(guid'" + accountId[0].id + "')";

    var data = Helper.Read(queryurl);
    if (data != null) {

        var address = "";
        if (data.Address1_StateOrProvince != null && data.Address1_StateOrProvince != "null" && data.Address1_StateOrProvince != "undefined") {
            address = data.Address1_StateOrProvince;
        }

        if (address + data.Address1_Line1 != null && address + data.Address1_Line1 != "null" && address + data.Address1_Line1 != "undefined") {
            address = address + data.Address1_City;
        }

        if (address + data.Address1_Line1 != null && address + data.Address1_Line1 != "null" && address + data.Address1_Line1 != "undefined") {
            address = address + data.Address1_Line1;
        }

        Xrm.Page.getAttribute("new_accountaddress").setValue(address);

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

