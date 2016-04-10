using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins.Common
{
    class QueryOwner
    {
        internal static EntityReference QueryPostSaleCsiOwner(Entity vehicle, IOrganizationService orgService)
        {
            EntityReference vehicleOwner = vehicle["new_vehicleowner"] as EntityReference;
            Entity account = orgService.Retrieve(vehicleOwner.LogicalName, vehicleOwner.Id, new ColumnSet(true));
            if (account != null)
            {
                if (account.Contains("new_province") == true)
                {
                    bool isInBeiJing = account.FormattedValues["new_province"] == "北京";
                    if (isInBeiJing == true)
                    {
                        EntityReference ownerTeam = new EntityReference();
                        ownerTeam.LogicalName = "team";
                        ownerTeam.Id = GetHeadQuarterCcrTeam(orgService);
                    }
                    else
                    {
                        EntityReference ownerUser = new EntityReference()
                        {
                            LogicalName = "systemuser",
                            Id = GetLocalCcrUser(account.FormattedValues["new_province"], orgService)
                        };
                    }

                }
                else
                {
                    throw new Exception(String.Format("车主{0}资料中，缺少省份信息，无法判断其归属地", account.Id));
                }
            }
            throw new Exception("未找到主键为{0}的客户记录");
        }

        private static Guid GetLocalCcrUser(string provinceName, IOrganizationService orgService)
        {
            //根据省份信息，获取其所属分公司的CCR用户
            throw new NotImplementedException();
        }

        private static Guid GetHeadQuarterCcrTeam(IOrganizationService orgService)
        {
            //获取总部的呼叫中心团队ID
            throw new NotImplementedException();
        }
    }
}
