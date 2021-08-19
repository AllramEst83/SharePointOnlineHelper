using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using AppForSharePointOnlineWebToolkit;


namespace SharePointOnlineHelper
{
    /*
     * FÖR TILLFÄLLET SKAPA EN MYSESSIONS KLASS OCH OCH KLISTRA IN KODEN I DITT PROJEKT
     * 
     * MySession must capture spContext in Index.
      var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
      MySession.Current.SpContext = spContext;
    */

    //public class MySession
    //{
    //    // private constructor
    //    public MySession() { }

    //    // Gets the current session.
    //    public static MySession Current
    //    {
    //        get
    //        {
    //            MySession session =
    //                (MySession)HttpContext.Current.Session["__MySession__"];
    //            if (session == null)
    //            {
    //                session = new MySession();
    //                HttpContext.Current.Session["__MySession__"] = session;
    //            }
    //            return session;

    //        }

    //    }
    //    // **** add your session properties here, e.g like this:
    //    public SharePointContext SpContext { get; set; }
    //}
}
