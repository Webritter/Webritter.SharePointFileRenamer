using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using System.Collections;

namespace Webritter.SharePointFileRenamer
{
    class Program
    {

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            log.Info("Programm started ");
            if (args.Count() != 1)
            {
                Console.WriteLine("Missing Parameter: optionXmlFile");
                log.Error("Missing Parameter: optionXmlFile");
                return;
            }

            if (args[0] == "sample")
            {
                RunOptions.GreateSampleXml("sample.xml");
                string message = "Sample xml file created: sample.xml";
                Console.WriteLine(message);
                log.Info(message);
                return;

            }



            string xmlFileName = args[0];
            if (!System.IO.File.Exists(xmlFileName))
            {
                Console.WriteLine("File does not exist: " + xmlFileName);
                log.Error("File does not exist: " + xmlFileName);
                return;
            }

            RunOptions options;
            try
            {
                options = RunOptions.LoadFromXMl(xmlFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Can't read and validate xmlFile: " + xmlFileName);
                log.Error("Can't read and validate xmlFile: " + xmlFileName);
                return;
            }


            if (string.IsNullOrEmpty(options.SiteUrl))
            {
                string message = "Missing SiteUrl in xmlFile: " + xmlFileName;
                Console.WriteLine(message);
                log.Error(message);
                return;
            }

            //Get instance of Authentication Manager  
            AuthenticationManager authenticationManager = new AuthenticationManager();
            //Create authentication array for site url,User Name,Password and Domain  
            try
            {
                SecureString password = GetSecureString(options.Password);
                //Create the client context  
                using (var ctx = authenticationManager.GetNetworkCredentialAuthenticatedContext(options.SiteUrl, options.Username, password, options.Domain))
                //using (var ctx = authenticationManager.GetWebLoginClientContext(options.SiteUrl))
                {
                    Site site = ctx.Site;
                    ctx.Load(site);
                    ctx.ExecuteQuery();
                    log.Info("Succesfully authenticated as " + options.Username);

                    List spList = ctx.Web.Lists.GetByTitle(options.LibraryName);
                    ctx.Load(spList);
                    ctx.ExecuteQuery();

                    log.Info("DocumentLibrary found with total " + spList.ItemCount + " docuements");

                    if (spList != null && spList.ItemCount > 0)
                    {
                        // build viewFields
                        string viewFields;
                        viewFields = "<ViewFields>";
       
                        foreach(var field in options.FieldNames)
                        {
                            viewFields += "<FieldRef " +
                                "Name='" + field.FieldName + "'" +
                                ((field.IsLookup) ? " LookupId='True' Type='Lookup' " : "") +
                                " />";
                        }
                        viewFields += "</ViewFields>";

                        // build caml query
                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View> " +
                                                "<Query>" +
                                                    options.CamlQuery +
                                                "</Query>" +
                                                viewFields +
                                            "</View>";

                        ListItemCollection listItems = spList.GetItems(camlQuery);
                        ctx.Load(listItems);
                        ctx.ExecuteQuery();

                        log.Info("found " +listItems.Count + " documents to check");


                        foreach (var item in listItems)
                        {
                            log.Info("Checking '" + item["FileLeafRef"] + "' ....");
                            bool skip = false;
                            List<object> fieldValues = new List<object>();
                            foreach(var field in options.FieldNames)
                            {
                                if (item[field.FieldName] == null)
                                {
                                    // the content of the field is null
                                    if (field.ShouldNotBeNull)
                                    {
                                        skip = true;
                                        log.Warn("Skipped because '" + field.FieldName + "' is null");
                                        break;
                                    }
                                    fieldValues.Add("");
                                }
                                else
                                {
                                    if (IsDictionary(item[field.FieldName]))
                                    {
                                        // this field is a managed metadata
                                        Dictionary<string,object> value = (dynamic)item[field.FieldName];
                                        if (value.ContainsKey("Label"))
                                        {
                                            fieldValues.Add(value["Label"]);
                                        }
                                        else
                                        {

                                        }

                                    }
                                    else
                                    {
                                        fieldValues.Add(item[field.FieldName]);
                                    }
                                    
                                }
                            }
                            if (!skip)
                            {
                                try
                                {
                                    string oldFileName = item["FileLeafRef"].ToString();
                                    string newFileName = string.Format(options.FileNameFormat, fieldValues.ToArray());
                                    item["FileLeafRef"] = newFileName;
                                    item.Update();
                                    ctx.ExecuteQuery();
                                    log.Info("Renamed ''" + oldFileName + "' to '" + newFileName +"'");
                                }
                                catch (Exception ex)
                                {
                                    log.Error("Exception renaing file: '" + listItems[0]["FileLeafRef"] + "'", ex);
                                }
                            }
                        }
                        
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception : " + ex.Message);
            }


            return;




            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            using (var context = new ClientContext(options.SiteUrl))
            {
                if (!string.IsNullOrEmpty(options.Username) && !string.IsNullOrEmpty(options.Password))
                {
                    var passWord = new SecureString();
                    foreach (char c in options.Password.ToCharArray()) passWord.AppendChar(c);
                    context.Credentials = new SharePointOnlineCredentials(options.Username, passWord);
                }




                Web web = context.Web;

                context.Load(web);
                context.ExecuteQuery();


                
            }

        }

        private static SecureString GetSecureString(string pwd)
        {
            var passWord = new SecureString();
            foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
            return passWord;
        }

        private static bool IsDictionary(object o)
        {
            if (o == null) return false;
            return o is IDictionary &&
                   o.GetType().IsGenericType &&
                   o.GetType().GetGenericTypeDefinition().IsAssignableFrom(typeof(Dictionary<,>));
        }

    }
}
