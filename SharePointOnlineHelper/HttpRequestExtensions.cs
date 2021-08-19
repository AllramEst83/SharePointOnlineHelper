using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Routing;

namespace SharePointOnlineHelper
{
    public static class HttpRequestExtensions
    {

        //used to get necc sp-addin-parameters when redirecting 
        public static RouteValueDictionary GetSpAddinParameters(this HttpRequestBase pRequest, Dictionary<string, string> extraParams)
        {
            var rvs = new RouteValueDictionary
            {
                {AddinQString.SP_HOST_URL,  pRequest[AddinQString.SP_HOST_URL] },
                {AddinQString.SP_PRODUCT_NUMBER ,pRequest[AddinQString.SP_PRODUCT_NUMBER]},
                {AddinQString.SP_APP_WEB_URL, pRequest[AddinQString.SP_APP_WEB_URL]},
                {AddinQString.SP_LANGUAGE , pRequest[AddinQString.SP_LANGUAGE]},
                {AddinQString.SP_CLIENT_TAG , pRequest[AddinQString.SP_CLIENT_TAG] },
                { AddinFormParams.SP_APP_TOKEN , pRequest[AddinFormParams.SP_APP_TOKEN] }
            };


            foreach (var extra in extraParams)
            {
                rvs.Add(extra.Key, extra.Value);
            }

            return rvs;
        }


        //used to get necc sp-addin-parameters when redirecting 
        public static object GetSpAddinParameters(this HttpRequestBase pRequest)
        {
            return new
            {
                SPHostUrl = pRequest[AddinQString.SP_HOST_URL],
                SPProductNumber = pRequest[AddinQString.SP_PRODUCT_NUMBER],
                SPAppWebUrl = pRequest[AddinQString.SP_APP_WEB_URL],
                SPLanguage = pRequest[AddinQString.SP_LANGUAGE],
                SPClientTag = pRequest[AddinQString.SP_CLIENT_TAG],
                SPAppToken = pRequest[AddinFormParams.SP_APP_TOKEN]
            };
        }


        //public static class RequestExtensions
        //{
        //Use this class like this:
        //var qString = this.Request.GetSpAddinParameters();
        //return RedirectToAction("Index", qString);
        //----------------------------------------------
        public static object GetSpAddinParameters2(this HttpRequestBase pRequest)
        {
            return new
            {
                SPHostUrl = pRequest["SPHostUrl"],
                SPProductNumber = pRequest["SPProductNumber"],
                SPAppWebUrl = pRequest["SPAppWebUrl"],
                SPLanguage = pRequest["SPLanguage"],
                SPClientTag = pRequest["SPClientTag"],
                SPAppToken = pRequest["SPAppToken"],
                Redirecting = "1"
            };
        }



        //     @{
        //      var url = Url.Action("CloseProject", "Home") + "?id={{item.ID}}";
        //      url = HttpUtility.UrlDecode(url);
        //      }
        //      data-ng-href="@url"

        //      Guid id = new Guid(Request["id"]);
    }
}
}
