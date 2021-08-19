using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AppForSharePointOnlineWebToolkit;
using Microsoft.SharePoint.Client;
using SharePointOnlineHelper.Model;

namespace SharePointOnlineHelper
{
    public class SharePointCrudOperations
    {
        //SUBSITES-------------------------------------------------
        //---------------------------------------------------------

        //Optional parameters
        public List<Web> GetSubsites(SharePointContext spContext)
        {
            //MySession must capture spContext in Index.
            //if spContext is not passed as a parameter.
            //So it can be used here properly.
            //DONT FORGET TO GIVE THE APP APP CALL PERSMISSIONS AND MANAGE LISTS OR SITEASSETS UNDER PERMISSIONS

            ErrorHelper errorHelper = new ErrorHelper();
            List<Web> subSiteList = new List<Web>();

            if (spContext == null) return subSiteList;
            using (var context = spContext.CreateAppOnlyClientContextForSPHost())
            {
                var web = context.Web;
                try
                {
                    context.Load(web, website => website.Webs, website => website.Title);
                    context.ExecuteQuery();
                }
                catch (Exception e)
                {
                    errorHelper.Info(e.Message);
                }
                foreach (var subWeb in web.Webs)
                {
                    if (subWeb.WebTemplate != "APP")
                    {
                        subSiteList.Add(subWeb);
                    }
                }
            }
            return subSiteList;
        }


        public void CreateSubSite(SharePointContext spContext = null, string title = "", string description = "")
        {
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    if (title == "")
                    {
                        title = "New Subsite";
                    }
                    else if (description == "")
                    {
                        description = "No description";
                    }

                    WebCreationInformation creation = new WebCreationInformation
                    {
                        Title = title,
                        Description = description,
                        Url = title
                    };
                    context.Web.Webs.Add(creation);

                    //To retrive
                    //Web newWeb = context.Web.Webs.Add(creation);

                    //    context.Load(newWeb, w => w.Title);
                    //    context.ExecuteQuery();
                }
            }

        }
        //DeleteSubSite
        public void DeleteSubSite(SharePointContext spContext)
        {
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    context.Web.DeleteObject();

                }
            }
        }

        public void UpdateSubSiteProperties(string description = "", string url = "", string title = "", SharePointContext spContext = null)
        {
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    Web web = context.Web;
                    if (!String.IsNullOrWhiteSpace(title))
                    {

                        web.Title = title;
                    }
                    else if (!String.IsNullOrEmpty(description))
                    {
                        web.Description = description;
                    }
                    web.Update();
                    context.ExecuteQuery();

                }
            }


        }

        //SUBSITES----END------------------------------------------
        //---------------------------------------------------------


        //LISTS---------------------------------------------------
        //--------------------------------------------------------

        //RetriveListFromSharePoint

        public ListItemCollection RetriveList(ClientContext context, string listTitle)
        {
            //Hämtar SharePoint listan "Holidays"
            List getHolidayDates = context.Web.Lists.GetByTitle(listTitle);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();

            //Kan också specificera sin förfrågan:

            //            var camlQueryStr = "<View><Query><Where><Contains><FieldRef
            //Name = 'Company' ></ FieldRef >< Value Type = 'Text' > " + mySearchString +
            //"</Value></Contains></Where></Query></View>”;

            //query.ViewXml = camlQueryStr;

            ListItemCollection items = getHolidayDates.GetItems(query);

            context.Load(items);
            context.ExecuteQuery();

            return items;
        }

        //RetriveListFromSharePoint

        //RetriveAllSharePointListsInAWeb
        public List<List> RetriveAllSharePointListsInAWeb(SharePointContext spContext)
        {
            List<List> listOfWebs = new List<List>();
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    Web web = context.Web;
                    context.Load(web.Lists, lists => lists.Include(list => list.Title, list => list.Id));
                    context.ExecuteQuery();

                    foreach (List list in web.Lists)
                    {
                        listOfWebs.Add(list);
                    }
                }
            }

            return listOfWebs;
        }
        //CreateGenericList
        public void CreateGenericList(SharePointContext spContext = null, string title = "", string description = "")
        {
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    Web web = context.Web;

                    if (title == "")
                    {
                        title = "List with no title";
                    }

                    ListCreationInformation creationInfo = new ListCreationInformation
                    {
                        Title = title,
                        TemplateType = (int)ListTemplateType.GenericList

                    };

                    if (description == "")
                    {
                        description = "No description";
                    }
                    List list = web.Lists.Add(creationInfo);
                    list.Description = description;

                    list.Update();
                    context.ExecuteQuery();
                }
            }


        }
        //DeleteList
        public void DeleteList(SharePointContext spContext = null, string listTitle = "")
        {
            if (spContext != null && listTitle != "")
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {

                    List list = context.Web.Lists.GetByTitle(listTitle);
                    list.DeleteObject();

                    context.ExecuteQuery();
                }
            }

        }
        //GetListItems
        public List<ListItem> GetListItems(SharePointContext spContext = null, string title = "")
        {
            List<ListItem> listItems = new List<ListItem>();
            if (spContext != null && title != "")
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {

                    List itemsToGet = context.Web.Lists.GetByTitle(title);

                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection items = itemsToGet.GetItems(query);

                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem listItem in items)
                    {
                        listItems.Add(listItem);
                    }
                }
            }

            return listItems;
        }
        //AddListItemToList
        public void AddListItemToList(SharePointContext spContext, string listName, List<AddListItemToListModel> model)
        {
            List<Field> fieldList = new List<Field>();
            FieldNumber fldNumber = null;
            using (var context = spContext.CreateAppOnlyClientContextForSPHost())
            {
                var list = context.Web.Lists.GetByTitle(listName);
                foreach (var item in model)
                {
                    Field field = null;
                    switch (item.Type)
                    {
                        case "DateTime":
                            field = list.Fields.AddFieldAsXml(
                                "<Field Name =" + item.Name + " DisplayName=" + item.DisplayName + " Type=" +
                                item.Type + "  Format=" + item.Format + "/>", true, AddFieldOptions.DefaultValue);
                            break;

                        case "Text":
                            field = list.Fields.AddFieldAsXml(
                                 "<Field Name =" + item.Name + " DisplayName=" + item.DisplayName + " Type=" +
                                 item.Type + "/>", true, AddFieldOptions.DefaultValue);
                            break;

                        case "Choice":
                            field = list.Fields.AddFieldAsXml(
                                "<Field Name =" + item.Name + " DisplayName=" + item.DisplayName + " Type=" +
                                item.Type +
                                "<CHOICES>" +
                                " <CHOICE>" + item.ChoiceOne + "</CHOICE>" +
                                " <CHOICE>" + item.Choicetwo + "</CHOICE>" +
                                " <CHOICE>" + item.ChoiceThree + "</CHOICE>" +
                                " </CHOICES>/>", true, AddFieldOptions.DefaultValue);
                            break;

                        case "Currency":
                            field = list.Fields.AddFieldAsXml(
                                "<Field Name =" + item.Name + " DisplayName=" + item.DisplayName + " Type=" +
                                item.Type + "Decimals = " + item.Decimals + "Min = " + item.Min + "/>", true, AddFieldOptions.DefaultValue);
                            break;

                        case "Number":
                            field = list.Fields.AddFieldAsXml(
                                "<Field Name =" + item.Name + " DisplayName=" + item.DisplayName + " Type=" +
                                item.Type + "/>", true, AddFieldOptions.DefaultValue);
                            fldNumber = context.CastTo<FieldNumber>(field);

                            fldNumber.MaximumValue = item.NumMax;
                            fldNumber.MinimumValue = item.NumMin;
                            fldNumber.Update();

                            break;
                    }
                    fieldList.Add(field);
                }

                foreach (var fields in fieldList)
                {
                    fields.Update();
                }
                context.ExecuteQuery();

            }
        }
        //UpdateListItem
        public void UpdateListItem(SharePointContext spContext, string listTitle, int listItemId, string listItemName, string itemContent)
        {
            using (var context = spContext.CreateAppOnlyClientContextForSPHost())
            {
                List list = context.Web.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(listItemId);
                //Hitta de egentliga columnnamnen i url:en i List Settings
                listItem[listItemName] = itemContent;
                listItem.Update();

                context.ExecuteQuery();
            }
        }
        //DeleteListItem
        public void DeleteListItem(SharePointContext spContext, string listName, int listItemId)
        {
            using (var context = spContext.CreateAppOnlyClientContextForSPHost())
            {
                List announcementsList = context.Web.Lists.GetByTitle(listName);
                ListItem listItem = announcementsList.GetItemById(listItemId);
                listItem.DeleteObject();

                context.ExecuteQuery();
            }
        }
        //GetAllFieldsInlist
        public List<Field> GetAllFieldsInList(SharePointContext spContext, string listTitle)
        {
            List<Field> fieldlist = new List<Field>();
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    List list = context.Web.Lists.GetByTitle(listTitle);
                    context.Load(list.Fields);

                    // We must call ExecuteQuery before enumerate list.Fields. 
                    context.ExecuteQuery();

                    foreach (Field field in list.Fields)
                    {
                        fieldlist.Add(field);
                    }

                }
            }
            return fieldlist;
        }
        //GetSpecificFieldFromList
        public FieldText GetSpecificFieldFromList(SharePointContext spContext, string listTitle, string fieldTitle)
        {
            FieldText textField = null;
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    List list = context.Web.Lists.GetByTitle(listTitle);
                    Field field = list.Fields.GetByInternalNameOrTitle(fieldTitle);
                    textField = context.CastTo<FieldText>(field);

                    context.Load(textField);
                    context.ExecuteQuery();

                }
            }


            return textField;
        }


        //LISTS------END------------------------------------------
        //--------------------------------------------------------


        //GROUPS--------------------------------------------------
        //--------------------------------------------------------

        //AddUserToGroup
        public void AddUserToGroup(SharePointContext spContext, string logInname, string eMail, string title, string groupName)
        {
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {

                    GroupCollection siteGroups = context.Web.SiteGroups;

                    //Group membersGroup = siteGroups.GetById(5);
                    Group membersGroup = siteGroups.GetByName(groupName);

                    // Let's set up the new user info. 
                    UserCreationInformation userCreationInfo = new UserCreationInformation
                    {
                        Email = eMail,//"user@domain.com",
                        LoginName = logInname,//"domain\\user",
                        Title = title//"Mr User"
                    };

                    membersGroup.Users.Add(userCreationInfo);
                    context.ExecuteQuery();

                }
            }

        }
        //Retrieve all users in a SharePoint group
        public List<User> GetallUsersFromGroup(SharePointContext spContext, string groupTitle)
        {
            List<User> userList = new List<User>();
            if (spContext != null)
            {
                using (var context = spContext.CreateAppOnlyClientContextForSPHost())
                {

                    GroupCollection siteGroups = context.Web.SiteGroups;

                    //Group membersGroup = siteGroups.GetById(5);
                    Group membersGroup = siteGroups.GetByName(groupTitle);
                    context.Load(membersGroup.Users);
                    context.ExecuteQuery();

                    foreach (var user in membersGroup.Users)
                    {
                        userList.Add(user);
                    }

                }
            }
            return userList;
        }
        //

        //GROUPS----END-------------------------------------------
        //UseFullOperations--------------------------------------
        ////ConvertEnumToString
        public String convertToString(this Enum eff)
        {
            return Enum.GetName(eff.GetType(), eff);
        }
        //ConvertEnumToString
        //ConvertStringToEnum
        public EnumType converToEnum<EnumType>(this String enumValue)
        {
            return (EnumType)Enum.Parse(typeof(EnumType), enumValue);
        }
        //ConvertStringToEnum
        //UseFullOperations-----END----------------------------------
        //--------------------------------------------------------
    }
}
