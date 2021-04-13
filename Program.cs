using CredentialManagement;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ClassicPageContentReplacement
{
    class Program
    {
        static void Main(string[] args)
        {
            TestClassicPage();
        }
        private static ClientContext GetClientContext(string url, string credentialStoreName = "SPO-M365x725618")
        {
            SecureString se = new SecureString();
            Credential cred = new Credential() { Target = "SPO-M365x725618" };
            ClientContext ret = new ClientContext(url);
            cred.Load();
            ret.Credentials = new SharePointOnlineCredentials(cred.Username, cred.SecurePassword);
            return ret;
        }
        private static void TestClassicPage()
        {
            string url = "https://m365x725618.sharepoint.com/sites/ClassicPublishing01";
            string searchKeyword = "Contoso";
            string replaceKeyword = "NewContoso";
            using (ClientContext context = GetClientContext(url))
            {
                // enumaret "Pages" Lib
                var files = context.Web.Lists.GetByTitle("Pages").RootFolder.Files;
                context.Load(files,
                    fs => fs.Include(f => f.ServerRelativeUrl),
                    fs => fs.Include(f => f.CheckedOutByUser.UserPrincipalName),
                    fs => fs.Include(f => f.ListItemAllFields),
                    fs => fs.Include(f => f.CheckOutType));
                context.ExecuteQuery();

                foreach (var file in files)
                {
                    if (file.CheckOutType == CheckOutType.None)
                    {
                        //string pageUrl = "https://m365x725618.sharepoint.com/sites/ClassicPublishing01/Pages/TestClassicPage.aspx";
                        //Microsoft.SharePoint.Client.File file = context.Web.GetFileByUrl(pageUrl);

                        LimitedWebPartManager manager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                        context.Load(manager.WebParts,
                            wps => wps.Include(wp => wp.WebPart.Title),
                            wps => wps.Include(wp => wp.WebPart.ZoneIndex),
                            wps => wps.Include(wp => wp.WebPart.Properties),
                            wps => wps.Include(wp => wp.Id),
                            wps => wps.Include(wp => wp.ZoneId));
                        context.ExecuteQuery();

                        //check out file for change. 
                        bool pageChanged = false;
                        file.CheckOut();

                        //Update PublishingPageContent
                        string pageContent = file.ListItemAllFields["PublishingPageContent"].ToString();
                        if (pageContent.IndexOf(searchKeyword, StringComparison.CurrentCultureIgnoreCase) != -1)
                        {
                            file.ListItemAllFields["PublishingPageContent"] = pageContent.Replace(searchKeyword, replaceKeyword);
                            file.ListItemAllFields.Update();
                            pageChanged = true;
                        }

                        foreach (var wp in manager.WebParts)
                        {
                            if (wp.WebPart.Title == "Content Editor")
                            {
                                //export the web to get CEWP's content
                                var wpXml = manager.ExportWebPart(wp.Id);
                                context.ExecuteQuery();

                                string wpContent = wpXml.Value;

                                if (wpContent.IndexOf(searchKeyword, StringComparison.CurrentCultureIgnoreCase) != -1)
                                {
                                    Console.WriteLine("Found '{0}' at {1}", searchKeyword, file.ServerRelativeUrl);
                                    string zoneId = wp.ZoneId;
                                    int zoneIndex = wp.WebPart.ZoneIndex;

                                    //replace the content
                                    string wpNewContext = wpContent.Replace(searchKeyword, replaceKeyword);

                                    //remove old wp
                                    wp.DeleteWebPart();
                                    context.ExecuteQuery();

                                    //import the wp
                                    var wpDef = manager.ImportWebPart(wpNewContext);
                                    manager.AddWebPart(wpDef.WebPart, zoneId, zoneIndex);
                                    context.ExecuteQuery();
                                    Console.WriteLine("Replace '{0}' to '{1}' at {2}", searchKeyword, replaceKeyword, file.ServerRelativeUrl);

                                    pageChanged = true;
                                }
                            }

                        }
                        if (pageChanged)
                        {
                            file.CheckIn("replace content", CheckinType.MajorCheckIn);
                        }
                        else
                        {
                            file.UndoCheckOut();
                        }
                        context.ExecuteQuery();
                        Console.WriteLine("File '{0}' was processed and {1} changes", file.ServerRelativeUrl, pageChanged ? "found" : "no");
                    }
                    else
                    {
                        Console.WriteLine("File '{0}' was checked out by {1}", file.ServerRelativeUrl, file.CheckedOutByUser.UserPrincipalName);
                    }
                }
            }

        }
    }
}
