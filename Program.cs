using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.IO;
using System.Security;
using System.Threading;
using Microsoft.SharePoint.Client;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            string usernameDonMT = "spoammgd3";
            string passwordDonMT = "Dried0806";
            string usernameMT = "spoam";
            string passwordMT = "E^spo@mSPM";
            HttpWebRequest req;

#if NOTYET
            /* Access MT with hostname */
            req = CreateWebRequest("134.170.210.58",
                "spoam2bd989c30c3c4739988426cb5ed541df.sharepoint.com",
                usernameMT,
                passwordMT,
                true /*useMySite*/,
                false /* useIP */);
            ValidateWebRequest(req);

            /* Access MT with IP */
            req = CreateWebRequest("134.170.210.58",
                "spoam2bd989c30c3c4739988426cb5ed541df.sharepoint.com",
                usernameMT,
                passwordMT,
                true /*useMySite*/,
                true /* useIP */);
            ValidateWebRequest(req);

            /* Access DonMT with hostname */
            req = CreateWebRequest("104.146.0.104",
                "spoam1d3329c5f3c0421eac9203d6b2f52a93.035dapp.com",
                usernameDonMT,
                passwordDonMT,
                true /*useMySite*/,
                false /*useIP */);
            ValidateWebRequest(req);

            /* Access DonMT with IP */
            req = CreateWebRequest("104.146.0.104",
                "spoam1d3329c5f3c0421eac9203d6b2f52a93.035dapp.com",
                usernameDonMT,
                passwordDonMT,
                true /*useMySite*/,
                true /*useIP*/);
#endif

            req = CreateWebRequest("104.146.0.104",
                "spoamed7439374b9043af8e54555aa9a526bb-my.035dapp.com",
                usernameDonMT,
                passwordDonMT,
                true /*useMySite*/,
                false /*useIP*/);
            ValidateWebRequest(req);

            req = CreateWebRequest("104.146.0.104",
                "spoamed7439374b9043af8e54555aa9a526bb-my.035dapp.com",
                usernameDonMT,
                passwordDonMT,
                true /*useMySite*/,
                true /*useIP*/);
            ValidateWebRequest(req);
        }

        public const int CookieExpireTime = 4;
        private const string CookieStringFormat = "{0}={1};";
        static public HttpWebRequest CreateWebRequest(string hostVip, string hostName, string username, string password, bool useMySite, bool useIP)
        {
            Uri uri;
            string liveid = String.Format("{0}@spotxmon.ccsctp.net", username);
            SecureString securePassword = new SecureString();
            Array.ForEach(password.ToCharArray(), c => securePassword.AppendChar(c));

            if (useMySite)
            {
                uri = new Uri(String.Format("https://{0}/personal/{1}_spotxmon_ccsctp_net/Documents/Forms/All.aspx", 
                        useIP ? hostVip : hostName, username));
            }
            else
            {
                uri = new Uri(String.Format("https://{0}/_layouts/15/start.aspx#/SitePages/DevHome.aspx",
                        useIP ? hostVip : hostName));
            }

            System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(uri);
            Cookie[] cookies = GetCookieFromLiveID(uri, useIP ? hostName : null, liveid, securePassword);
            req.CookieContainer = new System.Net.CookieContainer();
            foreach (Cookie cookie in cookies)
            {
                cookie.Domain = hostName;
                req.CookieContainer.Add(cookie);
            }
            
            req.Method = WebRequestMethods.Http.Get;
            req.PreAuthenticate = true;
            req.Proxy = HttpWebRequest.GetSystemWebProxy();                 // this allows use of Fiddler to look at requests            
            req.AllowAutoRedirect = true;
            req.MaximumAutomaticRedirections = 300;
            req.Accept = "*/*";
            req.Headers.Add("Accept-Language", "en-US");
            req.KeepAlive = true;
            req.Timeout = 30000;
            if (useIP)
            {
                req.Host = hostName;
            }
            req.Headers.Add("Front-End-Https", "On");
            req.ContentType = "text/xml";
            //req.UserAgent = "Mozilla/4.0+(compatible;+MSIE+5.01;+Windows+NT+5.0";
            req.UseDefaultCredentials = false;
            req.Headers.Add("SPResponseGuid", Guid.NewGuid().ToString());
            req.Headers.Add("X-SPOSyntheticTransaction", "True");           // By adding this, response will have Server Side and IIS Latency as well as Hashed Machine name data
            req.Headers.Add("X-RequestForceAuthentication", "true");        //req.Headers.Add("Accept-Encoding", "gzip, deflate");
            req.Headers.Add("X-IDCRL_ACCEPTED", "t");

            return req;
        }

        static public void ValidateWebRequest(HttpWebRequest request)
        {
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Console.WriteLine("Content length is {0}", response.ContentLength);
            Console.WriteLine("Content type is {0}", response.ContentType);

            // Get the stream associated with the response.
            Stream receiveStream = response.GetResponseStream();

            // Pipes the stream to a higher level stream reader with the required encoding format. 
            StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);

            Console.WriteLine("Response stream received.");
            Console.WriteLine(readStream.ReadToEnd());
            response.Close();
            readStream.Close();
        }

        /// <summary>
        /// Validates the input cookies and return the valid ones.
        /// </summary>
        /// <param name="input">The input cookies array</param>
        /// <returns>The cookies array or NULL if there are no valid cookies found.</returns>
        internal static Cookie[] ValidateCookies(Cookie[] input)
        {
            if (input != null)
            {
                var validCookies = input.Where(cookie => !String.IsNullOrEmpty(cookie.Name) && !String.IsNullOrEmpty(cookie.Value));
                if (validCookies.Count() > 0)
                {
                    return validCookies.ToArray<Cookie>();
                }
            }

            return null;
        }

        /// <summary>
        /// Format the cookie string from a cookie object list
        /// </summary>
        /// <param name="cookieList">Cookie object list</param>
        /// <returns>Cookie string</returns>
        public static string CookieListToString(Cookie[] cookieList)
        {
            string cookieString = string.Empty;

            foreach (Cookie cookie in cookieList)
            {
                cookieString += string.Format(                    
                    CookieStringFormat,
                    cookie.Name,
                    cookie.Value);
            }

            return cookieString;
        }

        public static Cookie[] CreateCookieFromString(string cookieData, Uri siteUrl)
        {
            List<Cookie> returnCookies = new List<Cookie>();

            if (!String.IsNullOrEmpty(cookieData))
            {
                string[] cookies = cookieData.Split(';');
                if (cookies.Length == 0)
                {
                    return null;
                }

                string key, name;
                int firstIndex;
                foreach (string cookieString in cookies)
                {
                    if (String.IsNullOrEmpty(cookieString))
                    {
                        continue;
                    }
                    key = "";
                    name = "";
                    firstIndex = -1;

                    firstIndex = cookieString.IndexOf('=');
                    if (firstIndex <= -1)
                    {
                        continue;
                    }

                    key = cookieString.Substring(0, firstIndex);
                    name = cookieString.Substring(firstIndex + 1, cookieString.Length - firstIndex - 1);

                    if (String.IsNullOrEmpty(key) || String.IsNullOrEmpty(name))
                    {
                        continue;
                    }

                    returnCookies.Add(new Cookie(key, name, "/", siteUrl.GetComponents(UriComponents.Host, UriFormat.UriEscaped)));
                }
            }

            if (returnCookies.Count == 0)
            {
                return null;
            }

            return returnCookies.ToArray<Cookie>();
        }

        internal static Cookie[] CreateCookieFromStringIDCRL(string cookieData, Uri siteUrl)
        {
            if (String.IsNullOrEmpty(cookieData))
            {
                return new Cookie[] { };
            }

            Cookie[] returnValue = new Cookie[1];
            int index = cookieData.IndexOf('=');
            returnValue[0] = new Cookie(cookieData.Substring(0, index), cookieData.Substring(index + 1), "/", siteUrl.GetComponents(UriComponents.Host, UriFormat.UriEscaped));

            return returnValue;
        }

        internal static Cookie[] GetCookieFromLiveID(Uri siteUrl, string hostHeader, string liveId, SecureString livePassword)
        {
            Cookie[] cookies = null;

            // For primary probes, siteUrl is like https://hostname/suffix and hostHeader is null;
            // For DR probes, siteUrl is like https://IP/suffix and hostHeader is hostname
            cookies = GetCookiesUsingSharePointOnlineCredentials(liveId, livePassword, siteUrl, hostHeader);

            //Apply the expire time workaround to keep the cookie for a long time
            if (cookies != null)
            {
                foreach (Cookie cookie in cookies)
                {
                    if (cookie.Expires <= cookie.TimeStamp)
                    {
                        cookie.Expires = cookie.TimeStamp.AddHours(CookieExpireTime);
                    }
                }
            }

            return cookies;
        }

        /// <summary>
        /// Get cookies using SharePoint Online credentials
        /// </summary>
        /// <param name="username">user name</param>
        /// <param name="password">password</param>
        /// <param name="siteUrl">site url this cookie is for</param>
        /// <param name="hostHeader">hostHeader string</param>
        /// <param name="ParentProbe">ParentProbe: use it to output log message, if null, no log message will be written</param>
        /// <returns></returns>
        public static Cookie[] GetCookiesUsingSharePointOnlineCredentials(string username, SecureString password, Uri siteUrl, string hostHeader)
        {
            string cookieStr = null;
            int retryInterval = 1;
            if (!string.IsNullOrEmpty(hostHeader))
            {
                // DR SubsSettings DB might be restoring a tlog when probe tries to get cookie. Restoring takes 10+ secs so making retry interval 5 secs
                retryInterval = 5;
            }
            DoActionWithRetry(() =>
            {
                SharePointOnlineCredentials spCredentials = new SharePointOnlineCredentials(username, password);
                if (!string.IsNullOrEmpty(hostHeader))
                {
                    spCredentials.ExecutingWebRequest += delegate(object sender, SharePointOnlineCredentialsWebRequestEventArgs e)
                    {
                        if (e.WebRequest.RequestUri.GetLeftPart(UriPartial.Authority) == siteUrl.GetLeftPart(UriPartial.Authority))
                        {
                            e.WebRequest.Host = hostHeader;
                            e.WebRequest.Headers["Front-End-Https"] = "On";
                        }
                    };
                }

                try
                {
                    cookieStr = spCredentials.GetAuthenticationCookie(url: siteUrl, alwaysThrowOnFailure: true);
                }
                catch (WebException webException)
                {
                    throw webException;
                }

                if (cookieStr != null)
                {
                    cookieStr = cookieStr.Trim();
                }

                if (String.IsNullOrEmpty(cookieStr))
                {
                    throw new NullReferenceException("Empty/Null cookies retrieved from OrgId");
                }
            }
                , 3 /* retry */
                , retryInterval /* wait sec in between */
                , null);

            Cookie[] cookies = CreateCookieFromStringIDCRL(cookieStr, siteUrl);
            cookies = ValidateCookies(cookies);
            return cookies;
        }

        /// <summary>
        /// Helper function to perform a retry action
        /// </summary>
        /// <param name="a">The acion</param>
        /// <param name="maxRetries">Max number of retries</param>
        /// <param name="waitBetweenRetrySec">Wait secs between retries</param>
        /// <param name="logs">Log string</param>
        private static void DoActionWithRetry(Action a, int maxRetries, int waitBetweenRetrySec, StringBuilder logs)
        {
            if (a == null)
            {
                throw new ArgumentNullException("No action specified");
            }

            do
            {
                try
                {
                    a();
                    break;
                }
                catch (Exception ex)
                {
                    if (maxRetries <= 0)
                    {
                        throw;
                    }
                    else
                    {
                        Thread.Sleep(TimeSpan.FromSeconds(waitBetweenRetrySec));
                    }
                }
            } while (maxRetries-- > 0);
        }
    }
}
