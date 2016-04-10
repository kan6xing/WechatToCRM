using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.GoldenHarvest.Plugins
{
    public class AccountBirthdaySetPlugin : IPlugin
    {

        private const string C_EntityName = "account";
        private const string C_ImageName = "Image";

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService orgService = factory.CreateOrganizationService(null);

                if (ValidInput(context) == false)
                {
                    return;
                }

                if (context.MessageName == "Create")
                {
                    DoCreate(context, orgService);
                }
                else if (context.MessageName == "Update")
                {
                    DoUpdate(context, orgService);
                }
            }
            catch (FaultException<OrganizationServiceFault> excp)
            {
                throw new InvalidPluginExecutionException(excp.Message);
            }
            catch (Exception excp)
            {
                throw new InvalidPluginExecutionException(excp.Message);
            }
        }

        private void DoUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            DateTime? preBirthday = GetBirthday(preImage);
            DateTime? postBirthday = GetBirthday(postImage);

            if (preBirthday == postBirthday)
            {
                return;
            }

            if (postBirthday.HasValue == false)
            {
                Entity acc = new Entity(context.PrimaryEntityName);
                acc.Id = context.PrimaryEntityId;
                acc["new_nextbirthday"] = null;

                orgService.Update(acc);
            }
            else
            {
                Entity acc = new Entity(context.PrimaryEntityName);
                acc.Id = context.PrimaryEntityId;
                acc["new_nextbirthday"] = CalcNextBirthday(postBirthday.Value);

                orgService.Update(acc);
            }
        }

        private DateTime? GetBirthday(Entity account)
        {
            if (account.Contains("new_birthday") == false)
            {
                return null;
            }
            else
            {
                return (DateTime)account["new_birthday"];
            }
        }

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity acc = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
            if (acc.Contains("new_birthday") == false)
            {
                return;
            }

            DateTime birthday = ((DateTime)acc["new_birthday"]);
            DateTime nextBirthday = CalcNextBirthday(birthday);

            acc = new Entity(context.PrimaryEntityName);
            acc.Id = context.PrimaryEntityId;
            acc["new_nextbirthday"] = nextBirthday;

            orgService.Update(acc);
        }

        private static DateTime CalcNextBirthday(DateTime birthday)
        {
            birthday = birthday.ToLocalTime();
            DateTime nextBirthday;
            DateTime birthdayThisYear = new DateTime(DateTime.Today.Year, birthday.Month, birthday.Day);
            if (birthdayThisYear <= DateTime.Today)
            {
                nextBirthday = birthdayThisYear.AddYears(1);
            }
            else
            {
                nextBirthday = birthdayThisYear;
            }
            return nextBirthday;
        }

        private bool ValidInput(IPluginExecutionContext context)
        {
            if (context.PrimaryEntityName != C_EntityName)
            {
                return false;
            }

            if (context.Stage != 40)
            {
                return false;
            }

            if (context.Depth > 1)
            {
                return false;
            }

            if (context.MessageName != "Update" &&
                context.MessageName != "Create")
            {
                return false;
            }

            if (context.MessageName == "Update")
            {
                if (context.PreEntityImages.ContainsKey(C_ImageName) == false ||
               context.PostEntityImages.ContainsKey(C_ImageName) == false)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
