using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SharePointOnlineHelper
{
    public class ErrorHelper
    {

        public HttpException ThrowCustomException(int statusCode, string customMessage)
        {

            switch (statusCode)
            {
                case 404:
                    return new HttpException(404, "Not found: " + customMessage);
                case 500:
                    return new HttpException(500, "Server error: " + customMessage);
                default:
                    return new HttpException(400, customMessage);
            }
        }

        public void Info(string pMessage)
        {
            Trace.TraceError($"{pMessage}");
        }

        public void Error(string pMessage, Exception pException)
        {
            Trace.TraceError($"{pMessage}, Message: {pException.Message}, StackTrace:{pException.StackTrace}");
        }

    }
}
