function orderDetail_onLoad() {


    //调用设置字段操作权限方法。
    //Authority();

    //选定定价方式：调用选定定价方式方法，选定“自定义价格”或“系统定价”，并将其置为不可编辑
    PricingDisplay();

    //“错误提示”字段的显示，如果此字段有值则显示，否则不显示“售价”字段的显示，如果此字段有值则显示，否则不显示.调用显示产品内容方法，显示各产品的特定字段
    isDisplay();

    //“售价”字段的显示，如果此字段有值则显示，否则不显示
    if (Xrm.Page.getAttribute("new_closeprice").getValue() == null) {
        Xrm.Page.getControl("new_closeprice").setVisible(false);
    }

    //调用显示产品内容方法，显示各产品的特定字段

    //Add by Changjun.Yao at 2013-11-19 用于用于有效期控制
    GetCinemaID();
    GetValidDates();
    //End of adding by Changjun.Yao at 2013-11-19 用于用于有效期控制


    //在应收净额字段不为空时，显示售价、基础价
    //if(Xrm.Page.getAttribute("baseamount").getValue()!=null){
    //Xrm.Page.getControl("new_closeprice").setVisible(true);
    // Xrm.Page.getControl("new_baseprice").setVisible(true);
    //}

}

