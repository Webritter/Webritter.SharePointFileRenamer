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
using Microsoft.SharePoint.Workflow;

namespace Webritter.SharePointFileRenamer
{
    class Program
    {

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            log.Info("Programm started ");
            #region check parameter file
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
                return;
            }

            if (string.IsNullOrEmpty(options.SiteUrl))
            {
                string message = "Missing SiteUrl in xmlFile: " + xmlFileName;
                Console.WriteLine(message);
                log.Error(message);
                return;
            }
            #endregion
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

                    WorkflowRepository workflowRep = null;
                    // check if there is a workflow to start
                    if (!string.IsNullOrEmpty(options.WorkflowName)) 
                    {
                        workflowRep = new WorkflowRepository(ctx);
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
                            #region get field values
                            List<object> fieldValues = new List<object>();
                            foreach (var field in options.FieldNames)
                            {
                                object fieldValue = GetFieldValueAsIntOrString(item, field.FieldName);                                                             
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
                            #endregion
                            string currentFileName = null;
                            string newFileName = null;
                            if (!skip)
                            {
                                try
                                {
                                    // find current and new file name
                                    currentFileName = item["FileLeafRef"].ToString();
                                    string extension = System.IO.Path.GetExtension(currentFileName);
                                    newFileName = string.Format(options.FileNameFormat, fieldValues.ToArray()) + extension;
                                    // replace illegal characters
                                    newFileName = string.Join("_", newFileName.Split(System.IO.Path.GetInvalidFileNameChars()));

                                    // get current and new path
                                    ctx.Load(item.File);
                                    ctx.ExecuteQuery();
                                    string currentPath = item.File.ServerRelativeUrl.ToString().Replace("/" + currentFileName, "");
                                    string newPath = currentPath;
                                    if (!string.IsNullOrEmpty(options.MoveTo))
                                    {
                                        if (!options.MoveTo.StartsWith("/"))
                                        {
                                            // new path is a sub folder
                                            newPath = currentPath + "/" + options.MoveTo;
                                        }
                                        else
                                        {
                                            // new path is server relative
                                            newPath = options.MoveTo;
                                        }
                                    }

                                    // get current and new status
                                    string currentStatus = null;
                                    string newStatus = null;
                                    if(!string.IsNullOrEmpty(options.StatusFieldName))
                                    {
                                        var val = GetFieldValueAsIntOrString(item, options.StatusFieldName);
                                        currentStatus = (val != null) ? val.ToString() : "";
                                        newStatus = options.StatusSuccessValue;
                                    }



                                    if (newFileName != currentFileName || newPath != currentPath || newStatus != currentStatus)
                                    {
                                        // there s something todo with the file!! 

                                        try
                                        {
                                            bool hasCheckedOut = false;
                                            #region checkout

                                            if (spList.ForceCheckout)
                                            {
                                                try
                                                {
                                                    item.File.CheckOut();
                                                    ctx.ExecuteQuery();
                                                    hasCheckedOut = true;
                                                    log.Info("Checked out: ''" + currentFileName + "'");
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
                                                    log.Info("File was checked out by me: ''" + currentFileName + "'");
                                                }
                                            }

                                            #endregion

                                            #region rename or move file
                                            if (newPath != currentPath)
                                            {
                                                // move the file
                                                item.File.MoveTo(newPath + "/" + newFileName, MoveOperations.Overwrite);
                                                ctx.ExecuteQuery();
                                                log.Info("Moved ''" + item["FileLeafRef"] + "' to '" + options.MoveTo + "'");
                                            }
                                            else if (newFileName != currentFileName)
                                            {
                                                // rename the file
                                                item["FileLeafRef"] = newFileName;
                                                item.Update();
                                                ctx.ExecuteQuery();
                                                log.Info("Renamed '" + currentFileName + "' to '" + newFileName + "'");

                                            }
                                
                                            #endregion
                                            #region update status

                                            if (newStatus != currentStatus)
                                            {
                                                // update status in item
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
                                                        termValues = item[options.StatusFieldName] as TaxonomyFieldValueCollection;
                                                        bool found = false;
                                                        foreach (TaxonomyFieldValue tv in termValues)
                                                        {
                                                            found = found || tv.TermGuid == statusTermId;
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
                                                        termValue.Label = newStatus;
                                                        termValue.TermGuid = statusTermId;
                                                        termValue.WssId = -1;
                                                        txField.SetFieldValueByValue(item, termValue);
                                                    }
                                                }
                                                else
                                                {
                                                    // try to setup status as text field 
                                                    item[options.StatusFieldName] = newStatus;
                                                }
                                                item.Update();
                                                ctx.ExecuteQuery();
                                                log.Info("Updated:'" + options.StatusFieldName + "' from '" + currentStatus +  "' to '" + newStatus + "'");



                                            }
                                            #endregion
                                            #region update version
                                            #endregion

                                            if (hasCheckedOut)
                                            {
                                                #region checkin
                                                if (hasCheckedOut)
                                                {
                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    realoadedItem.File.CheckIn(options.CheckinMessage, options.CheckinType);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Checked in: ''" + newFileName + "'");
                                                }
                                                #endregion
                                                #region publish
                                                if (!string.IsNullOrEmpty(options.PublishInfo))
                                                {
                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    realoadedItem.File.Publish(options.PublishInfo);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Published: ''" + newFileName + "'");
                                                }
                                                #endregion
                                                #region approve
                                                if (!string.IsNullOrEmpty(options.ApproveInfo))
                                                {
                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    realoadedItem.File.Approve(options.ApproveInfo);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Approved: ''" + newFileName + "'");
                                                }
                                                #endregion
                                                #region workflow
                                                if (!string.IsNullOrEmpty(options.WorkflowName))
                                                {

                                                   

                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    ctx.Load(realoadedItem);
                                                    ctx.ExecuteQuery();
                                                    workflowRep.RunListWorkflow45(realoadedItem.Id, options.WorkflowName, null);

                                                }
                                                #endregion
                                            }



                                        }
                                        catch (Exception ex)
                                        {
                                            log.Error("Error renaiming or moving file", ex);
                                        }


                                    }
                                    else
                                    {
                                        log.Info("Skipped because filename or path is not to change!");
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

        private static object GetFieldValueAsIntOrString(ListItem item, string fieldName)
        {
            object fieldValue = null;
            if (item[fieldName] != null)
            {
                // try to get a good field value
                if (item[fieldName] is TaxonomyFieldValue)
                {
                    // this field is a single taxonomy field
                    var taxValue = TaxonomyHelpers.GetTaxonomyFieldValue(item, fieldName);
                    fieldValue = taxValue.Label;
                }
                else if (item[fieldName] is TaxonomyFieldValueCollection)
                {
                    // this field is a multi taxonomy field
                    var taxValues = TaxonomyHelpers.GetTaxonomyFieldValueCollection(item, fieldName);
                    if (taxValues.Count > 0)
                    {
                        // join all labels with comma
                        fieldValue = string.Join(",", taxValues.Select(v => v.Label));
                    }
                }
                else if (item[fieldName] is FieldUrlValue)
                {
                    var urlValue = item[fieldName] as FieldUrlValue;
                    fieldValue = urlValue.Description;
                }
                else if (item[fieldName] is FieldUserValue)
                {
                    var usrValue = item[fieldName] as FieldUserValue;
                    fieldValue = usrValue.LookupValue;
                }
                else if (IsDictionary(item[fieldName]))
                {
                    // other field types not supportd yet
                    log.Error("This field type is not supported yet!!");

                }
                else
                {
                    fieldValue = item[fieldName];
                }

            }
            return fieldValue;
        }









    }



}
