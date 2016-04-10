var serverUrl = Xrm.Page.context.getServerUrl();
var ODataPath = serverUrl + "/XRMServices/2011/OrganizationData.svc";
var productType;//产品类型
var Filedstate;//字段状态
var rolesName;//角色名称

//Add by Changjun.Yao at 2013-11-19  用于控制有效期
var cinemaid;
var validdates = 0;
//End of adding by Changjun.Yao at 2013-11-19  用于控制有效期

//var ownerid;//负责人id
//var ordertype;//订单类型
//var examineandverifystate;//激活审核状态
//var gatheringreviewstatus;//收款审核状态
//var userid=Xrm.Page.context.getUserId().toLowerCase();//;//当前登录用户id

function PricingDisplay() //控制选定定价方式
{
    Xrm.Page.ui.controls.get("ispriceoverridden").setDisabled(true);
    //var salesorderdetailid=Xrm.Page.data.entity.getId();//订单产品id
    //salesorderInfo(salesorderdetailid);//订单id
    // ownerInfo(salesorderid);//获取负责人id
    //var fromtype=Xrm.Page.ui.getFormType();//获取当前窗体状态

    if (Xrm.Page.getAttribute("productid").getValue() != null) {
        var productid = Xrm.Page.getAttribute("productid").getValue()[0].id;
        productInfo(productid); //获取产品类型
        isDisplay(); //显示产品内容
        if (productType == 100000000 || productType == 100000001 || productType == 100000004 || productType == 100000002 || productType == 100000005 || productType == 100000003 || productType == 100000006 || productType == 100000011) //设置系统定价与自定义定价
        {
            Xrm.Page.getAttribute("ispriceoverridden").setValue(false);
            Xrm.Page.ui.controls.get("priceperunit").setDisabled(true);
        } else {
            Xrm.Page.getAttribute("ispriceoverridden").setValue(true);
            Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        }
        //if(fromtype==2)
        //{
        //   Authority();//设置字段操作权限
        // }
    }
}

function productInfo(productid) //使用ajax获取产品类型
{
    var productInfoReq = new XMLHttpRequest();
    productInfoReq.open("GET", ODataPath + "/ProductSet(guid'" + productid + "')", false);
    productInfoReq.setRequestHeader("Accept", "application/json");
    productInfoReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    productInfoReq.onreadystatechange = function () { productInfoReqCallBack(this); };
    productInfoReq.send();
}

function productInfoReqCallBack(productInfoReq) {
    if (productInfoReq.readyState == 4) {
        if (productInfoReq.status == 200) {
            var productInfo = JSON.parse(productInfoReq.responseText).d;
            productType = productInfo.new_ProductType.Value;
        }
    }
}

function salesorderInfo(salesorderdetailid) {
    var salesorderInfoReq = new XMLHttpRequest();
    salesorderInfoReq.open("GET", ODataPath + "/SalesOrderDetailSet(guid'" + salesorderdetailid + "')", false);
    salesorderInfoReq.setRequestHeader("Accept", "application/json");
    salesorderInfoReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    salesorderInfoReq.onreadystatechange = function () { var salesorderid = salesorderInfoReqCallBack(this); };
    salesorderInfoReq.send();
}

function salesorderInfoReqCallBack(salesorderInfoReq) {
    if (salesorderInfoReq.readyState == 4) {
        if (salesorderInfoReq.status == 200) {
            var salesorderInfo = JSON.parse(salesorderInfoReq.responseText).d;
            salesorderid = salesorderInfo.SalesOrderId.Id;
        }
    }
}

function rolesType(id) {
    var rolesReq = new XMLHttpRequest();
    rolesReq.open("GET", ODataPath + "/RoleSet(guid'" + id + "')", false);
    rolesReq.setRequestHeader("Accept", "application/json");
    rolesReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    rolesReq.onreadystatechange = function () { rolesReqCallBack(this); };
    rolesReq.send();
}

function rolesReqCallBack(rolesReq) {
    if (rolesReq.readyState == 4) {
        if (rolesReq.status == 200) {
            var rolesInfo = JSON.parse(rolesReq.responseText).d;
            rolesName = rolesInfo.Name;
            //ordertype=ownerInfo.new_OrderType.Value;
            //examineandverifystate=ownerInfo.new_examineandverifystate.Value;
            //gatheringreviewstatus=ownerInfo.new_gatheringreviewstatus.Value;
        }
    }
}

function isDisplay() //根据产品类型显示相应产品内容
{
    if (productType == 100000005) {
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(true);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(true);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(true);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(true);
        Xrm.Page.ui.controls.get("new_seating").setVisible(true);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setDisabled(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(true);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000002) {
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(true);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(true);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(true);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(true);
        Xrm.Page.ui.controls.get("new_seating").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(true);
        //Xrm.Page.ui.controls.get("new_cardnum").setVisible(true);
        //Xrm.Page.ui.controls.get("new_cardtimes").setVisible(true);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(true);
        //Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        //Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        //Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        //Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setDisabled(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("new_film").setVisible(true);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000001) {
        //var new_cardnum=Xrm.Page.getAttribute("new_cardnum").getValue();
        //var new_cardtimes=Xrm.Page.getAttribute("new_cardtimes").getValue();
        //Xrm.Page.getAttribute("quantity").setValue(new_cardnum*new_cardtimes);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(true);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(true);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(true);

        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setDisabled(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26

    } else if (productType == 100000000 || productType == 100000004 || productType == 100000003) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(true);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setDisabled(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000007) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(false);
        //Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000008) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        //Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000006) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(true);
        Xrm.Page.ui.controls.get("new_baseprice").setDisabled(true);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(true);
        //Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    } else if (productType == 100000010) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        //Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26

    } else if (productType == 100000009) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        //Xrm.Page.ui.controls.get("priceperunit").setDisabled(false);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26

    } else if (productType == 100000011) {
        Xrm.Page.ui.controls.get("new_baseprice").setVisible(false);
        Xrm.Page.ui.controls.get("new_closeprice").setVisible(false);
        Xrm.Page.ui.controls.get("priceperunit").setDisabled(true);
        Xrm.Page.ui.controls.get("new_showingstandnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_showingstandtime").setVisible(false);
        Xrm.Page.ui.controls.get("new_otherrights").setVisible(false);
        Xrm.Page.ui.controls.get("new_filmloungeid").setVisible(false);
        Xrm.Page.ui.controls.get("new_seating").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardnum").setVisible(false);
        Xrm.Page.ui.controls.get("new_cardtimes").setVisible(false);
        Xrm.Page.ui.controls.get("quantity").setDisabled(false);
        Xrm.Page.ui.controls.get("new_discountamount").setVisible(false);
        Xrm.Page.ui.controls.get("new_servicetime").setVisible(false);
        Xrm.Page.ui.controls.get("new_articleforsalecategory").setVisible(false);
        Xrm.Page.ui.controls.get("manualdiscountamount").setVisible(true);
        Xrm.Page.ui.controls.get("manualdiscountamount").setDisabled(true);
        Xrm.Page.ui.controls.get("new_film").setVisible(false);    //Add by Changjun.Yao at 2013/11/26
    }
}

function Authority() //控制订单编辑权限
{
    var aorderapprovestate = Xrm.Page.getAttribute("new_aorderapprovestate").getValue();//A单审核状态

    var borderfinapprovestate = Xrm.Page.getAttribute("new_borderfinapprovestate").getValue();//B单财务审核状态
    var bordermktapprovestate = Xrm.Page.getAttribute("new_bordermktapprovestate").getValue();//B单激活审核状态
    Xrm.Page.ui.controls.get("new_payway").setDisabled(true);//付款方式
    Xrm.Page.ui.controls.get("new_gatheringdate").setDisabled(true);//收款日期

    var userid = Xrm.Page.context.getUserId();//.toLowerCase();//当前登录用户id
    var ownerid = Xrm.Page.getAttribute("ownerid").getValue()[0].id;//获取负责人id
    var fromtype = Xrm.Page.ui.getFormType();//获取当前窗体状态
    var ordertype = Xrm.Page.getAttribute("new_ordertype").getValue();//获取订单类型  
    var roles = Xrm.Page.context.getUserRoles().toString();//获取用户角色ce5c4f5d-98f2-e211-b4ab-00259033ee4d 
    var rolesgroup = roles.split(",");
    //teamRoles(userid);
    var financeRoles;
    for (var i = 0; i <= rolesgroup.length; i++) {
        var rolesid = rolesgroup[i];
        rolesType(rolesid);
        if (rolesName == "影城财务")//判断是否存在财务角色
        {
            financeRoles = true;
            //alert("true");
        }
    }
    if (fromtype == 2) {
        if (ordertype == "100000000") {
            //alert(aorderapprovestate);
            //alert(userid);
            if (aorderapprovestate == 100000000 && ownerid == userid || aorderapprovestate == 100000006 && ownerid == userid) {
                return;
            } else if (financeRoles == true && aorderapprovestate == 100000001) {
                Xrm.Page.ui.controls.get("new_ordertype").setDisabled(true);
                Xrm.Page.ui.controls.get("new_proposalnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_saler").setDisabled(true);
                Xrm.Page.ui.controls.get("name").setDisabled(true);
                //Filedstate=true;
                Xrm.Page.ui.controls.get("customerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_accountaddress").setDisabled(true);
                Xrm.Page.ui.controls.get("new_contact").setDisabled(true);
                Xrm.Page.ui.controls.get("new_phonenumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_orderdate").setDisabled(true);
                Xrm.Page.ui.controls.get("pricelevelid").setDisabled(true);
                Xrm.Page.ui.controls.get("transactioncurrencyid").setDisabled(true);
                //Xrm.Page.ui.controls.get("new_isuptostandard").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinema").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaaccountnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaopeningbank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemabank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaphone").setDisabled(true);
                Xrm.Page.ui.controls.get("ownerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_gatheringdate").setDisabled(false);//付款日期
                Xrm.Page.ui.controls.get("new_payway").setDisabled(false);//付款方式
                Xrm.Page.ui.controls.get("new_isgoal").setDisabled(true);
                Xrm.Page.ui.controls.get("new_receivabledate").setDisabled(true);

            } else {
                Xrm.Page.ui.controls.get("new_ordertype").setDisabled(true);
                Xrm.Page.ui.controls.get("new_proposalnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_saler").setDisabled(true);
                //Xrm.Page.ui.controls.get("name").setDisabled(true);
                Filedstate = true;
                Xrm.Page.ui.controls.get("customerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_accountaddress").setDisabled(true);
                Xrm.Page.ui.controls.get("new_contact").setDisabled(true);
                Xrm.Page.ui.controls.get("new_phonenumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_orderdate").setDisabled(true);
                Xrm.Page.ui.controls.get("pricelevelid").setDisabled(true);
                Xrm.Page.ui.controls.get("transactioncurrencyid").setDisabled(true);
                //Xrm.Page.ui.controls.get("new_isuptostandard").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinema").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaaccountnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaopeningbank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemabank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaphone").setDisabled(true);
                Xrm.Page.ui.controls.get("ownerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_gatheringdate").setDisabled(true);//付款日期
                Xrm.Page.ui.controls.get("new_payway").setDisabled(true);//付款方式
                Xrm.Page.ui.controls.get("new_isgoal").setDisabled(true);
                Xrm.Page.ui.controls.get("new_receivabledate").setDisabled(true);
            }

        } else if (ordertype == "100000001") {
            if ((borderfinapprovestate == 100000000 && ownerid == userid || borderfinapprovestate == 100000005 && ownerid == userid) && (bordermktapprovestate == 100000000 && ownerid == userid || bordermktapprovestate == 100000003 && ownerid == userid)) {
                return;
            } else if (financeRoles == true && borderfinapprovestate == 100000001)
                //|| financeRoles==true && bordermktapprovestate==100000001 
            {
                Xrm.Page.ui.controls.get("new_ordertype").setDisabled(true);
                Xrm.Page.ui.controls.get("new_proposalnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_saler").setDisabled(true);
                Xrm.Page.ui.controls.get("name").setDisabled(true);
                Xrm.Page.ui.controls.get("customerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_accountaddress").setDisabled(true);
                Xrm.Page.ui.controls.get("new_contact").setDisabled(true);
                Xrm.Page.ui.controls.get("new_phonenumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_orderdate").setDisabled(true);
                Xrm.Page.ui.controls.get("pricelevelid").setDisabled(true);
                Xrm.Page.ui.controls.get("transactioncurrencyid").setDisabled(true);
                //Xrm.Page.ui.controls.get("new_isuptostandard").setDisabled(true);
                Xrm.Page.ui.controls.get("new_gatheringdate").setDisabled(false);//付款日期
                Xrm.Page.ui.controls.get("new_payway").setDisabled(false);//付款方式
                Xrm.Page.ui.controls.get("new_cinema").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaaccountnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaopeningbank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemabank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaphone").setDisabled(true);
                Xrm.Page.ui.controls.get("ownerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_isgoal").setDisabled(true);
                Xrm.Page.ui.controls.get("new_receivabledate").setDisabled(true);
            } else {
                Xrm.Page.ui.controls.get("new_ordertype").setDisabled(true);
                Xrm.Page.ui.controls.get("new_proposalnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_saler").setDisabled(true);
                //Xrm.Page.ui.controls.get("name").setDisabled(true);
                Filedstate = true;
                Xrm.Page.ui.controls.get("customerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_accountaddress").setDisabled(true);
                Xrm.Page.ui.controls.get("new_contact").setDisabled(true);
                Xrm.Page.ui.controls.get("new_phonenumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_orderdate").setDisabled(true);
                Xrm.Page.ui.controls.get("pricelevelid").setDisabled(true);
                Xrm.Page.ui.controls.get("transactioncurrencyid").setDisabled(true);
                //Xrm.Page.ui.controls.get("new_isuptostandard").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinema").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaaccountnumber").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaopeningbank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemabank").setDisabled(true);
                Xrm.Page.ui.controls.get("new_cinemaphone").setDisabled(true);
                Xrm.Page.ui.controls.get("ownerid").setDisabled(true);
                Xrm.Page.ui.controls.get("new_gatheringdate").setDisabled(true);//付款日期
                Xrm.Page.ui.controls.get("new_payway").setDisabled(true);//付款方式
                Xrm.Page.ui.controls.get("new_isgoal").setDisabled(true);
                Xrm.Page.ui.controls.get("new_receivabledate").setDisabled(true);
            }
        }
    }
}
//更改字段权限状态
function AuthorityField() {
    if (Filedstate == true) {
        Xrm.Page.ui.controls.get("name").setDisabled(true);
    }
}

//数量字段
function numberFiled(ExecutionObj) {
    Xrm.Page.ui.controls.get("quantity").setDisabled(false);
    var quantity = Xrm.Page.getAttribute("quantity").getValue();
    if (productType == 100000001) {
        if (quantity < 50) {
            //Xrm.Page.getAttribute("quantity").setValue(50);
            alert("数量不允许小于50");
            //ExecutionObj.getEventArgs().preventDefault();
        }
    } else if (productType != 100000007 && productType != 100000009 && productType != 100000010 && productType != 100000011 && productType != 100000002 && productType != 100000005) {
        if (quantity < 30) {
            //Xrm.Page.getAttribute("quantity").setValue(30);
            alert("数量不允许小于30");
            //ExecutionObj.getEventArgs().preventDefault();
        }
    }

    //Add by Changjun.Yao at 2013-11-19 验证有效期
    if (!IsDatesValid()) {
        alert("有效期不符合" + String(validdates) + "天的规定");
    }
    //End of Adding by Changjun.Yao at 2013-11-19 验证有效期
}

//是否是指定团票产品
function isAppointProduct() {
    var productid = Xrm.Page.getAttribute("productid").getValue()[0].id;
    productInfo(productid);
    if (productType == 100000000 || productType == 100000002 || productType == 100000005 || productType == 100000003 || productType == 100000004) {
        return true;
    } else {
        return false;
    }
}

function acountNumber() {
    var new_cardnum = Xrm.Page.getAttribute("new_cardnum").getValue();
    var new_cardtimes = Xrm.Page.getAttribute("new_cardtimes").getValue();
    Xrm.Page.getAttribute("quantity").setValue(new_cardnum * new_cardtimes);
}

function uomidInfo() //根据产品自动带出单位并控制产品内容显示
{
    var Isnull = Xrm.Page.getAttribute("productid").getValue();
    if (Isnull != null) {
        var productid = Xrm.Page.getAttribute("productid").getValue()[0].id;
        PricingDisplay();//控制产品内容显示    
        var uomidInfoReq = new XMLHttpRequest();
        uomidInfoReq.open("GET", ODataPath + "/ProductSet(guid'" + productid + "')", false);
        uomidInfoReq.setRequestHeader("Accept", "application/json");
        uomidInfoReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        uomidInfoReq.onreadystatechange = function () { uomidInfoReqCallBack(this); };
        uomidInfoReq.send();

        //Add by ChangjunYao at 2013-11-08
        if (productType == 100000011) {    //赠券产品
            var StartDate = new Date();
            var EndDate = AddMonth(StartDate, 3);
            if (EndDate != null) {
                Xrm.Page.getAttribute("new_startdate").setValue(StartDate);
                Xrm.Page.getAttribute("new_enddate").setValue(EndDate);
            }
        }
        //End of Adding by ChangjunYao at 2013-11-08
        GetValidDates(); //Add by ChangjunYao at 2013-11-19
    }
}

function uomidInfoReqCallBack(uomidInfoReq) {
    if (uomidInfoReq.readyState == 4) {
        if (uomidInfoReq.status == 200) {
            var uomidInfo = JSON.parse(uomidInfoReq.responseText).d;
            var uomidLookup = new Array();
            uomidLookup[0] = new Object();
            uomidLookup[0].entityType = "uom";
            uomidLookup[0].id = uomidInfo.DefaultUoMId.Id;
            uomidLookup[0].logicalName = uomidInfo.DefaultUoMId.LogicalName;
            uomidLookup[0].name = uomidInfo.DefaultUoMId.Name;
            Xrm.Page.getAttribute("uomid").setValue(uomidLookup);
        }
    }
}

//根据订单类型显示框架协议号
function changeProposalnumber() {
    var new_ordertype = Xrm.Page.getAttribute("new_ordertype").getValue();
    if (new_ordertype == 100000001) {
        Xrm.Page.ui.controls.get("new_proposalnumber").setVisible(true);
        Xrm.Page.ui.controls.get("new_receivabledate").setVisible(true);    //Add by ChangjunYao at 2013-11-08
    } else {
        Xrm.Page.ui.controls.get("new_proposalnumber").setVisible(false);
        Xrm.Page.ui.controls.get("new_receivabledate").setVisible(false);    //Add by ChangjunYao at 2013-11-08
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

//Add by ChangjunYao at 2013-11-08
//返回n月后的截止日期
function AddMonth(StartDate, MonthNumber) {
    return new Date(StartDate.getFullYear(), (StartDate.getMonth()) + MonthNumber, StartDate.getDate() - 1,
        StartDate.getHours(), StartDate.getMinutes(), StartDate.getSeconds());

    //var curMonth = StartDate.getMonth();
    //var endMonth = curMonth + MonthNumber;

    //if(endMonth 

}
//End of Adding by ChangjunYao at 2013-11-08

//Add by ChangjunYao at 2013-11-19 用于有效期控制
//获取影城ID
function GetCinemaID() {
    var Helper = new RESTHelper();

    var orderId = Xrm.Page.getAttribute("salesorderid").getValue();
    if (orderId == null) {
        alert("订单不能为空！");
        return;
    }
    //取订单上的影城
    var queryurl = serverUrl + "/XRMServices/2011/OrganizationData.svc/SalesOrderSet(guid'" +
      orderId + "')";
    var data = Helper.Read(queryurl);
    if (data != null) {
        cinemaid = data.new_Cinema.Id;
    }
}

//获取有效期天数
function GetValidDates() {
    if (Xrm.Page.getAttribute("productid").getValue() != null &&
        Xrm.Page.getAttribute("quantity").getValue() != null) {
        var productid = Xrm.Page.getAttribute("productid").getValue()[0].id;
        var quantity = Xrm.Page.getAttribute("quantity").getValue();
        var Helper = new RESTHelper();
        var queryurl = serverUrl + "/XRMServices/2011/OrganizationData.svc/new_validdatesSet?$filter=new_cinemaid/Id eq (guid'" +
            cinemaid + "') and new_productid/Id eq (guid'" + productid + "')" +
            " and new_startquantity le " + String(quantity) +
            " and new_endquantity ge " + String(quantity);
        var data = Helper.Read(queryurl);

        if (data.results.length > 0) {
            validdates = data.results[0].new_validdates;
        }
        else {
            validdates = 0;
        }
    }
    else {
        validdates = 0;
    }
}

function DateDiff(Big, Small) {
    return (Date.parse(Big.toString()) - Date.parse(Small.toString())) / 86400000;
}

//判断日期合法性
function IsDatesValid() {
    var StartDate=Xrm.Page.getAttribute("new_startdate").getValue();
    var EndDate = Xrm.Page.getAttribute("new_enddate").getValue();

    if (validdates == 0) {
        return true;
    }
    else {
        if (DateDiff(EndDate, StartDate) > validdates) {
            return false;
        }
        else {
            return true;
        }
    }
}
//End of Adding by ChangjunYao at 2013-11-19 用于有效期控制
