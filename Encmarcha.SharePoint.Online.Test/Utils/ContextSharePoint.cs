using System;
using System.Configuration;
using System.Security;
using Microsoft.SharePoint.Client;

namespace Encmarcha.SharePoint.Online.Test.Utils
{
    internal static class  ContextSharePoint
    {
        #region Properties
        private static string TenantUrl { get; }
        private static string UserName { get;  }
        private static SecureString Password { get;  }        
        #endregion

        #region Constructor
        static ContextSharePoint()
        {
            UserName = ConfigurationManager.AppSettings["OnlineUserName"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            Password = GetSecureString(password);            
            TenantUrl = ConfigurationManager.AppSettings["OnlineSiteCollection"];

        }
        #endregion

        #region Methods

        public static ClientContext CreateClientContext()
        {
            try
            {
                var credentials= new SharePointOnlineCredentials(UserName,Password);
                var context = new ClientContext(TenantUrl)
                {
                    Credentials = credentials
                };

                return context;

            }
            catch (Exception)
            {

                return null;
            }
        }

        private static SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString",
                    nameof(input));
            }

            var secureString = new SecureString();
            foreach (var c in input)
            {
                secureString.AppendChar(c);
            }
            return secureString;
        }
        public static bool VerifyServer(ClientContext site)
        {
            return site != null;
        }
        #endregion
    }
}
