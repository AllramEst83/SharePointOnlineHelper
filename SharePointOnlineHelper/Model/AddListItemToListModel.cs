using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SharePointOnlineHelper.Model
{
    public class AddListItemToListModel
    {

        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Type { get; set; }
        public string Format { get; set; }
        public string Decimals { get; set; }
        public string Min { get; set; }
        public int NumMax { get; set; }
        public int NumMin { get; set; }
        public string ChoiceOne { get; set; }
        public string Choicetwo { get; set; }
        public string ChoiceThree { get; set; }

    }
}
