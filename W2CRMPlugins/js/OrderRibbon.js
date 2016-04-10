/// <reference path="../JScript/XrmPageTemplate.js" />
function doAOrderSubmit() {
    var cur_userid = Xrm.Page.context.getUserId();//.toLowerCase();
    var salerid = Xrm.Page.getAttribute("new_saler").getValue()[0].id;

    if (cur_userid != salerid) {
        alert("只有销售员本人才能提交订单，无法提交");
        return;
    }

    var orderType = Xrm.Page.getAttribute("new_ordertype");
    if (orderType == null) {
        alert("缺少订单类型，无法提交");
        return false;
    }
    else {
        orderType = orderType.getValue();
    }

    if (orderType == 100000001)//B Order
    {
        alert("当前订单类型为B单，无法执行A单提交操作");
        return false;
    }

    var approveState = Xrm.Page.getAttribute("new_aorderapprovestate");
    if (approveState == null) {
        alert("缺少A单审批状态字段，无法提交");
        return false;
    }
    else {
        approveState.setValue(100000001);
    }

    Xrm.Page.data.entity.save();
}

function doBOrderSubmit_Fin() {
    var cur_userid = Xrm.Page.context.getUserId();//.toLowerCase();
    var salerid = Xrm.Page.getAttribute("new_saler").getValue()[0].id;

    if (cur_userid != salerid) {
        alert("只有销售员本人才能提交订单，无法提交");
        return;
    }

    var orderType = Xrm.Page.getAttribute("new_ordertype");
    if (orderType == null) {
        alert("缺少订单类型，无法提交");
        return false;
    }
    else {
        orderType = orderType.getValue();
    }

    if (orderType == 100000000)//A Order
    {
        alert("当前订单类型为A单，无法执行B单提交操作");
        return false;
    }

    var approveState = Xrm.Page.getAttribute("new_borderfinapprovestate");
    if (approveState == null) {
        alert("缺少B单财务审批状态字段，无法提交");
        return false;
    }
    else {
        approveState.setValue(100000001);
    }

    Xrm.Page.data.entity.save();
}

function doBOrderSubmit_Mkt() {
    var cur_userid = Xrm.Page.context.getUserId();//.toLowerCase();
    var salerid = Xrm.Page.getAttribute("new_saler").getValue()[0].id;

    if (cur_userid != salerid) {
        alert("只有销售员本人才能提交订单，无法提交");
        return;
    }

    var orderType = Xrm.Page.getAttribute("new_ordertype");
    if (orderType == null) {
        alert("缺少订单类型，无法提交");
        return false;
    }
    else {
        orderType = orderType.getValue();
    }

    if (orderType == 100000000)//A Order
    {
        alert("当前订单类型为A单，无法执行B单提交操作");
        return false;
    }

    var approveState = Xrm.Page.getAttribute("new_bordermktapprovestate");
    if (approveState == null) {
        alert("缺少B单市场审批状态字段，无法提交");
        return false;
    }
    else {
        approveState.setValue(100000001);
    }

    Xrm.Page.data.entity.save();
}