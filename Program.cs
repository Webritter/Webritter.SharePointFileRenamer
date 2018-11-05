using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using System.Collections;
using System.Collections.ObjectModel;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Globalization;

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

            if (!options.Enabled)
            {
                log.Info("Disabled by xml file");
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

                    Web web = ctx.Web;
                    ctx.Load(web);

                    User currentUser = web.CurrentUser;
                    ctx.Load(currentUser);

                    List spList = ctx.Web.Lists.GetByTitle(options.LibraryName);
                    ctx.Load(spList);

                    FieldCollection fields = ctx.Web.Fields;
                    ctx.Load(fields);

                    Field statusField = null;
                    if (!string.IsNullOrEmpty(options.StatusFieldName))
                    {
                        statusField = fields.GetByInternalNameOrTitle(options.StatusFieldName);
                        ctx.Load(statusField);
                    }
                    ctx.ExecuteQuery();

                    if (spList == null)
                    {
                        log.Error("List not found!");
                        return;
                    }

                    log.Info("Document library found with total " + spList.ItemCount + " docuements");

                    if (spList.ItemCount == 0)
                    {
                        return;
                    }
                    // check if the status field is a taxonomy field
                    // and try to find the correct term id in site collection
                    string statusTermId = null;
                    TaxonomyField txField = null;
                    if (statusField != null && statusField.TypeAsString == "TaxonomyFieldType")
                    {
                        txField = ctx.CastTo<TaxonomyField>(statusField);
                        statusTermId = TaxonomyHelpers.GetTermIdForTerm(options.StatusSuccessValue, txField.TermSetId, ctx);
                        log.Info("StatusField to update is a taxonomy field! id found:" + statusTermId);
                    }


                    // build viewFields
                    string viewFields;
                    viewFields = "<ViewFields>";
       
                    foreach(var field in options.FieldNames)
                    {
                        viewFields += "<FieldRef " +
                            "Name='" + field.FieldName + "'" +
                            " />";
                    }
                    if (!string.IsNullOrEmpty(options.StatusFieldName)) {
                        viewFields += "<FieldRef " +
                            "Name='" + options.StatusFieldName + "'" +
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
                        if (item.FileSystemObjectType == FileSystemObjectType.File)
                        {
                            List<object> fieldValues = new List<object>();
                            foreach (var field in options.FieldNames)
                            {
                                object fieldValue = null;
                                if (item[field.FieldName] != null) 
                                {
                                    // try to get a good field value
                                    if (item[field.FieldName] is TaxonomyFieldValue)
                                    {
                                        // this field is a single taxonomy field
                                        var taxValue = TaxonomyHelpers.GetTaxonomyFieldValue(item, field.FieldName);
                                        fieldValue = taxValue.Label;
                                    }
                                    else if (item[field.FieldName] is TaxonomyFieldValueCollection)
                                    {
                                        // this field is a multi taxonomy field
                                        var taxValues = TaxonomyHelpers.GetTaxonomyFieldValueCollection(item, field.FieldName);
                                        if (taxValues.Count > 0)
                                        {
                                            // join all labels with comma
                                            fieldValue = string.Join(",", taxValues.Select(v => v.Label));
                                        }
                                    }
                                    else if (item[field.FieldName] is FieldUrlValue)
                                    {
                                        var urlValue = item[field.FieldName] as FieldUrlValue;
                                        fieldValue = urlValue.Description;
                                    }
                                    else if (item[field.FieldName] is FieldUserValue)
                                    {
                                        var usrValue = item[field.FieldName] as FieldUserValue;
                                        fieldValue = usrValue.LookupValue;
                                    }
                                    else if (IsDictionary(item[field.FieldName]))
                                    {
                                        // other field types not supportd yet
                                        log.Error("This field type is not supported yet!!");
                                        return;
                                    }
                                    else
                                    {
                                        fieldValue = item[field.FieldName];
                                    }

                                }
                                fieldValues.Add(fieldValue);
                                if (fieldValue == null)
                                {
                                    // the content of the field is null
                                    if (field.ShouldNotBeNull)
                                    {
                                        skip = true;
                                        log.Warn("Skipped because '" + field.FieldName + "' is null");
                                        break;
                                    }
                                }
                            }
                            string fileName = null;
                            string newFileName = null;
                            if (!skip)
                            {
                                try
                                {
                                    // format new file name
                                    fileName = item["FileLeafRef"].ToString();
                                    string extension = System.IO.Path.GetExtension(fileName);
                                    newFileName = string.Format(options.FileNameFormat, fieldValues.ToArray()) + extension;
                                    // replace illegal characters
                                    newFileName = string.Join("_", newFileName.Split(System.IO.Path.GetInvalidFileNameChars()));
                                    if (newFileName != fileName)
                                    {
                                        #region checkout
                                        bool hasCheckedOut = false;
                                        if (spList.ForceCheckout)
                                        {
                                            try
                                            {
                                                item.File.CheckOut();
                                                ctx.ExecuteQuery();
                                                hasCheckedOut = true;
                                                log.Info("Checked out: ''" + fileName + "'");
                                            }
                                            catch (Exception ex)
                                            {
                                                // check if checked out by me
                                                User checkedOutBy = item.File.CheckedOutByUser;
                                                ctx.Load(checkedOutBy);
                                                ctx.ExecuteQuery();
                                                if (checkedOutBy.Id != currentUser.Id)
                                                {
                                                    log.Warn("Skipped because item is checked out by another user'");
                                                    continue;
                                                }
                                                // leave hasCheckedOu on false, because changes was made
                                                // by the sema user as script is running.
                                                //hasCheckedOut = true;
                                                log.Info("File was checked out by me: ''" + fileName + "'");
                                            }
                                        }
                                        #endregion
                                        #region rename file
                                        try
                                        {
                                            item["FileLeafRef"] = newFileName;
                                            item.Update();
                                            ctx.ExecuteQuery();
                                            log.Info("Renamed '" + fileName + "' to '" + newFileName + "'");
                                            fileName = newFileName;

                                            #endregion
                                            #region update status
                                            if (!string.IsNullOrEmpty(options.StatusFieldName))
                                            {

                                                if (statusTermId != null)
                                                {
                                                    // status field is a taxonomy field
                                                    // try to set the taxonomy field
                                                    TaxonomyFieldValue termValue = null;
                                                    TaxonomyFieldValueCollection termValues = null;

                                                    string termValueString = string.Empty;

                                                    if (txField.AllowMultipleValues)
                                                    {
                                                        // multi value taxonomy field as status field
                                                        // 
                                                        termValues = item[options.StatusFieldName] as TaxonomyFieldValueCollection;
                                                        bool found = false;
                                                        foreach (TaxonomyFieldValue tv in termValues)
                                                        {
                                                            found = found ||  tv.TermGuid == statusTermId;
                                                            termValueString += tv.WssId + ";#" + tv.Label + "|" + tv.TermGuid + ";#";
                                                        }

                                                        if (!found)
                                                        {
                                                            // add the new status at the end of all values
                                                            termValueString += "-1;#" + options.StatusSuccessValue + "|" + statusTermId;
                                                            termValues = new TaxonomyFieldValueCollection(ctx, termValueString, txField);
                                                            txField.SetFieldValueByValueCollection(item, termValues);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        // single value taxonomy field
                                                        termValue = new TaxonomyFieldValue();
                                                        termValue.Label = options.StatusSuccessValue;
                                                        termValue.TermGuid = statusTermId;
                                                        termValue.WssId = -1;
                                                        txField.SetFieldValueByValue(item, termValue);
                                                    }
                                                }
                                                else
                                                {
                                                    // try to setup status as text field 
                                                    item[options.StatusFieldName] = options.StatusSuccessValue;
                                                }
                                                // update status in item
                                                item.Update();
                                                ctx.ExecuteQuery();
                                                log.Info("status updated: ''" + options.StatusFieldName + "' to '" + options.StatusSuccessValue + "'");
                                            }
                                            #endregion
                                            #region checkin
                                            if (hasCheckedOut)
                                            {
                                                item.File.CheckIn(options.CheckinMessage, options.CheckinType);
                                                ctx.ExecuteQuery();
                                                log.Info("Checked in: ''" + newFileName + "'");
                                            }
                                            #endregion
                                            #region publish
                                            if (!string.IsNullOrEmpty(options.PublishInfo))
                                            {
                                                item.File.Publish(options.PublishInfo);
                                                ctx.ExecuteQuery();
                                                log.Info("Published: ''" + newFileName + "'");
                                            }
                                            #endregion
                                            #region approve
                                            if (!string.IsNullOrEmpty(options.ApproveInfo))
                                            {
                                                item.File.Approve(options.ApproveInfo);
                                                ctx.ExecuteQuery();
                                                log.Info("Approved: ''" + newFileName + "'");
                                            }
                                            #endregion
                                        }
                                        catch(Exception ex)
                                        {
                                            log.Error("Error renaiming File", ex);
                                        }
                                    } else
                                    {
                                        log.Info("Skipped because filename is not to change!");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error("Exception renaing file: '" + listItems[0]["FileLeafRef"] + "'", ex);
                                }
                            }
                             
                        }
                        else
                        {
                            // this is not a file
                            log.Warn("Skipped because this is not a file!");
                        }
                           
                    }
                        
                    

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception : " + ex.Message);
            }

            log.Info("Programm ended");

            return;





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
