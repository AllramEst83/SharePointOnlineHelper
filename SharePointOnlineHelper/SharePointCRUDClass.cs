using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineHelper
{
    class SharePointCRUDClass
    {

        //FRÅN Food International övningen - ServerTekniker

        private User _user;
        public string PRODUKTER { get; set; } = "Produkter";


        //RetriveListFromSharePoint

        public ListItemCollection RetriveList(ClientContext context, string listTitle)
        {
            //Hämtar SharePoint listan "Holidays"
            List retrivesList = context.Web.Lists.GetByTitle(listTitle);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = retrivesList.GetItems(query);

            context.Load(items);
            context.ExecuteQuery();

            return items;
        }

        //RetriveListFromSharePoint

        //GetSpecificListColum

        public List<Produkt> GetSpecificListColum(ClientContext context, string listTitle)
        {
            List<Produkt> produkter = new List<Produkt>();
            ListItemCollection listFromSharePoint = RetriveList(context, listTitle);
            foreach (ListItem items in listFromSharePoint)
            {

                Produkt oneItem = new Produkt
                {
                    id = (int)items["ID"],
                    produktNamn = (string)items["Title"],
                    leverantör = (string)items["_x0075_nf2"],
                    beskrivning = (string)items["wmsl"],
                    bild = ((FieldUrlValue)(items["Bild"])).Url
                };

                produkter.Add(oneItem);
            }

            return produkter;
        }

        //GetSpecificListColum

        //AddItemsToList


        public void AddItemsToList(Produkt formData, ClientContext context, string listTitle)
        {

            if (context != null)
            {
                List produktLista = context.Web.Lists.GetByTitle(listTitle);

                ListItemCreationInformation SVitemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = produktLista.AddItem(SVitemCreateInfo);

                newItem["Title"] = formData.produktNamn;
                newItem["_x0075_nf2"] = formData.leverantör;
                newItem["wmsl"] = formData.beskrivning;
                newItem["Bild"] = formData.bild;
                newItem["_x006f_f38"] = Enum.GetName(formData.stad.GetType(), formData.stad);

                newItem.Update();
                context.ExecuteQuery();

            }
        }


        //AddItemsToList

        //DeleteItem


        public void DeleteItem(ClientContext SPcontext, string listName, int listItemId)
        {
            if (SPcontext != null)
            {

                List listToDelete = SPcontext.Web.Lists.GetByTitle(listName);
                ListItem listItem = listToDelete.GetItemById(listItemId);
                listItem.DeleteObject();

                SPcontext.ExecuteQuery();

            }


        }


        //DeleteItem

        //ModifyItem


        public void ModifyItem(ClientContext context, string listTitle, Produkt formData)
        {

            List list = context.Web.Lists.GetByTitle(listTitle);
            ListItem newItem = list.GetItemById(formData.id);

            newItem["Title"] = formData.produktNamn;
            newItem["_x0075_nf2"] = formData.leverantör;
            newItem["wmsl"] = formData.beskrivning;
            newItem["Bild"] = formData.bild;
            newItem["_x006f_f38"] = Enum.GetName(formData.stad.GetType(), formData.stad);
            newItem.Update();

            context.ExecuteQuery();

        }


        //ModifyItem


        //GetSpecificItem
        public Produkt GetListItems(ClientContext context, int id, string listTitle)
        {
            Web web = context.Web;
            List list = web.Lists.GetByTitle(listTitle);
            var q = new CamlQuery() { ViewXml = $@"
                                                <View>
                                                    <Query>
                                                        <Where>
                                                           <Eq>
                                                            <FieldRef Name='ID' />
                                                             <Value Type='Counter'>{id}</Value>
                                                             </Eq>
                                                          </Where>
                                                      </Query>
                                                   </View>
                                                    " };
            var r = list.GetItems(q);
            context.Load(r);
            context.ExecuteQuery();

            var stad = (Stad)Enum.Parse(typeof(Stad), r[0]["_x006f_f38"].ToString());
            Produkt enProdukt = new Produkt
            {
                id = Convert.ToInt32(r[0]["ID"]),
                produktNamn = (string)r[0]["Title"],
                leverantör = (string)r[0]["_x0075_nf2"],
                beskrivning = (string)r[0]["wmsl"],
                stad = stad,
                bild = ((FieldUrlValue)(r[0]["Bild"])).Url

            };
            return enProdukt;
        }
        //GetSpecificItem

        //GetUserProperties
        public PersonProperties GetUserProperties(ClientContext context)
        {
            User _user = GetUser();

            PeopleManager peopleManager = new PeopleManager(context);
            PersonProperties personProperties = peopleManager.GetPropertiesFor(_user.LoginName);

            context.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
            context.ExecuteQuery();

            return personProperties;
        }
        //GetUserProperties

        //GetUser

        public User GetUser()
        {

            var spContext = MySession.Current.spcontext;

            using (var userContext = spContext.CreateUserClientContextForSPHost())
            {
                _user = userContext.Web.CurrentUser;
                try
                {
                    userContext.Load(_user);
                    userContext.ExecuteQuery();
                }
                catch (Exception e)
                {
                    throw new HttpException(400, "Error: " + e.Message);
                }
            }


            return _user;
        }

        //GetUser
    }
}
