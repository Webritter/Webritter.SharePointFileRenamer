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
using Microsoft.SharePoint.Client.WorkflowServices;
using System.Threading;

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
               log.Error("File does not exist: " + xmlFileName);
                return;
            }

            RunOptions runOptions;
            try
            {
                runOptions = RunOptions.LoadFromXMl(xmlFileName);
            }
            catch (Exception ex)
            {
                log.Error("Can't read and validate xmlFile: " + xmlFileName);
                return;
            }


            if (string.IsNullOrEmpty(runOptions.SiteUrl))
            {
                log.Error("Missing SiteUrl in xmlFile: " + xmlFileName);
                return;
            }
            #endregion

            //Get instance of Authentication Manager  
            AuthenticationManager authenticationManager = new AuthenticationManager();
            //Create authentication array for site url,User Name,Password and Domain  
            try
            {
                //SecureString password = GetSecureString(options.Password);
                //Create the client context  
                //using (var ctx = authenticationManager.GetNetworkCredentialAuthenticatedContext(options.SiteUrl, options.Username, password, options.Domain))
                using (var ctx = authenticationManager.GetWebLoginClientContext(runOptions.SiteUrl))
                {
                    Site site = ctx.Site;
                    ctx.Load(site);
                    ctx.ExecuteQuery();
                    log.Info("Succesfully authenticated as " + runOptions.Username);

                    Web web = ctx.Web;
                    ctx.Load(web);

                    var loopCnt = (runOptions.LoopCnt > 0) ? runOptions.LoopCnt : 1;
                    
                    while (loopCnt > 0)
                    {
                        foreach (var taskOptions in runOptions.Tasks)
                        {
                            if (!taskOptions.Enabled)
                            {
                                log.Info("Skipped task '" + taskOptions.Title + "' because task is disabled");
                                continue;
                            }

                            log.Info("Starting task '" + taskOptions.Title + "'");

                            // load the list of the document library
                            List spList = ctx.Web.Lists.GetByTitle(taskOptions.LibraryName);
                            ctx.Load(spList);


                            // loading all field definitions from sharepoint
                            FieldCollection fields = spList.Fields;
                            ctx.Load(fields);

                            // load the current user data
                            User currentUser = web.CurrentUser;
                            ctx.Load(currentUser);

                            // get all static uptade field values
                            List<FieldInfo> updateFields = new List<FieldInfo>();
                            foreach (var field in taskOptions.UpdateFields)
                            {
                                var fieldInfo = new FieldInfo();
                                fieldInfo.SpField = fields.GetByInternalNameOrTitle(field.FieldName);
                                // get the formating info for the new Value
                                fieldInfo.Format = field.Format;
                                fieldInfo.IsStaticValue = !field.Format.Contains("{");

                                ctx.Load(fieldInfo.SpField);
                                updateFields.Add(fieldInfo);
                            }
                            ctx.ExecuteQuery();

                            // Prepare static new values
                            foreach (var field in updateFields)
                            {
                                if (field.IsStaticValue)
                                {
                                    if (field.SpField != null && field.SpField.TypeAsString == "TaxonomyFieldType")
                                    {
                                        var txField = ctx.CastTo<TaxonomyField>(field.SpField);
                                        if (txField.AllowMultipleValues)
                                        {
                                            log.Error("update of multiple value taxonomy fields not supported jet");
                                            continue;
                                        }
                                        var statusTermId = TaxonomyHelpers.GetTermIdForTerm(field.Format, txField.TermSetId, ctx);
                                        log.Info("field to update is a taxonomy field! id found:" + statusTermId);
                                        if (statusTermId != null)
                                        {
                                            field.NewValue = new TaxonomyFieldValue()
                                            {
                                                Label = field.Format,
                                                TermGuid = statusTermId,
                                                WssId = -1
                                            };
                                            // new vaue to compare with current value as string is static 
                                            field.NewValueAsString = field.Format;
                                        }
                                        else
                                        {
                                            field.NewValueAsString = field.Format;
                                        }
                                    }
                                }
                                else
                                {
                                    if (field.SpField != null && field.SpField.TypeAsString == "TaxonomyFieldType")
                                    {
                                        log.Error("Dynamic field values in taxonomy field not supportede jet!");
                                        continue;
                                    }
                                }

                            }

                            log.Info("Document library found with total " + spList.ItemCount + " docuements");
                            if (spList.ItemCount == 0)
                            {
                                continue;
                            }

                            // build viewFields for CAML-Query with all query and update fields
                            string viewFields;
                            viewFields = "<ViewFields>";

                            foreach (var field in taskOptions.QueryFields)
                            {
                                viewFields += "<FieldRef " +
                                    "Name='" + field.FieldName + "'" +
                                    " />";
                            }
                            foreach (var field in taskOptions.UpdateFields)
                            {
                                viewFields += "<FieldRef " +
                                    "Name='" + field.FieldName + "'" +
                                    " />";
                            }

                            viewFields += "</ViewFields>";

                            // build caml query
                            string scope = (!string.IsNullOrEmpty(taskOptions.Scope)) ? " Scope='" + taskOptions.Scope + "' " : "";
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View " + scope + " > " +
                                                    "<Query>" +
                                                        taskOptions.CamlQuery +
                                                    "</Query>" +
                                                    viewFields +
                                                "</View>";

                            // get list items by CAML query
                            ListItemCollection listItems = spList.GetItems(camlQuery);
                            ctx.Load(listItems);
                            ctx.ExecuteQuery();

                            log.Info("found " + listItems.Count + " documents to check");



                            // go through all foind items
                            foreach (var item in listItems)
                            {
                                log.Info("Checking '" + item["FileLeafRef"] + "' ....");
                                bool skip = false;
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {
                                    #region get field values
                                    // get all query field values as string or integer
                                    // this values will be used to format the filename and the 
                                    // update fields
                                    List<object> fieldValues = new List<object>();
                                    foreach (var field in taskOptions.QueryFields)
                                    {
                                        object fieldValue = GetFieldValueAsIntOrString(item, field.FieldName);
                                        if (field.FieldName == "_UIVersionString")
                                        {
                                            // prepare the version value as a future major version
                                            // because it is impossible to update a major version!
                                            decimal dec;
                                            if (decimal.TryParse(fieldValue.ToString().Replace(".", ","), out dec))
                                            {
                                                fieldValue = Math.Ceiling(dec).ToString() + ".0";
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
                                    #endregion
                                    string currentFileName = currentFileName = item["FileLeafRef"].ToString(); ;
                                    string newFileName = currentFileName;
                                    if (!skip)
                                    {
                                        try
                                        {
                                            // get current values and new values of the fields to update
                                            bool anyFieldToChange = false;
                                            foreach (var field in updateFields)
                                            {
                                                field.CurrentValueAsString = GetFieldValueAsString(item, field.SpField.StaticName);
                                                if (!field.IsStaticValue)
                                                {
                                                    // this is a dynamic new value -> use format with query field values
                                                    field.NewValueAsString = string.Format(field.Format, fieldValues.ToArray());
                                                }
                                                anyFieldToChange = anyFieldToChange || field.CurrentValueAsString != field.NewValueAsString;
                                            }

                                            if (anyFieldToChange)
                                            {
                                                // there s something todo with the item!! 

                                                try
                                                {
                                                    bool hasCheckedOut = false;
                                                    #region checkout

                                                    if (spList.ForceCheckout || true)
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

                                                    #region update field
                                                    foreach (var field in updateFields)
                                                    {
                                                        if (field.CurrentValueAsString != field.NewValueAsString)
                                                        {
                                                            if (field.SpField.TypeAsString == "TaxonomyFieldType")
                                                            {
                                                                var txField = ctx.CastTo<TaxonomyField>(field.SpField);
                                                                txField.SetFieldValueByValue(item, field.NewValue);
                                                            }
                                                            else
                                                            {
                                                                item[field.SpField.InternalName] = field.NewValueAsString;
                                                            }
                                                            log.Info("Updated:'" + field.SpField.InternalName + "' from '" + field.CurrentValueAsString + "' to '" + field.NewValueAsString + "'");
                                                        }
                                                    }

                                                    item.Update();
                                                    ctx.ExecuteQuery();
                                                    #endregion

                                                    if (hasCheckedOut)
                                                    {
                                                        #region checkin
                                                        if (hasCheckedOut)
                                                        {
                                                            var realoadedItem = spList.GetItemById(item.Id);
                                                            realoadedItem.File.CheckIn(taskOptions.CheckinMessage, taskOptions.CheckinType);
                                                            ctx.ExecuteQuery();
                                                            log.Info("Checked in: '" + newFileName + "'");
                                                        }
                                                        #endregion
                                                        #region publish
                                                        if (!string.IsNullOrEmpty(taskOptions.PublishInfo))
                                                        {
                                                            var realoadedItem = spList.GetItemById(item.Id);
                                                            realoadedItem.File.Publish(taskOptions.PublishInfo);
                                                            ctx.ExecuteQuery();
                                                            log.Info("Published: '" + newFileName + "'");
                                                        }
                                                        #endregion
                                                        #region approve
                                                        if (!string.IsNullOrEmpty(taskOptions.ApproveInfo))
                                                        {
                                                            var realoadedItem = spList.GetItemById(item.Id);
                                                            realoadedItem.File.Approve(taskOptions.ApproveInfo);
                                                            ctx.ExecuteQuery();
                                                            log.Info("Approved: '" + newFileName + "'");
                                                        }
                                                        #endregion
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    log.Error("Error updating item", ex);
                                                }
                                            }
                                            else
                                            {
                                                log.Info("Skipped item update because nothing is to change!");
                                            }

                                            // find new file name
                                            if (!string.IsNullOrEmpty(taskOptions.FileNameFormat))
                                            {
                                                string extension = System.IO.Path.GetExtension(currentFileName);
                                                newFileName = string.Format(taskOptions.FileNameFormat, fieldValues.ToArray()) + extension;
                                                // replace illegal characters
                                                newFileName = string.Join("_", newFileName.Split(System.IO.Path.GetInvalidFileNameChars()));
                                            }

                                            string currentPath = "";
                                            string newPath = "";
                                            if (newFileName != currentFileName || !string.IsNullOrEmpty(taskOptions.MoveTo))
                                            {
                                                // get current and new path

                                                currentPath = item["FileRef"].ToString().Replace("/" + currentFileName, "");
                                                newPath = currentPath;
                                                if (!string.IsNullOrEmpty(taskOptions.MoveTo))
                                                {
                                                    newPath = taskOptions.MoveTo;

                                                    if (newPath.Contains("{"))
                                                    {
                                                        newPath = string.Format(newPath, fieldValues.ToArray());
                                                    }

                                                    if (!newPath.StartsWith("/"))
                                                    {
                                                        // new path is a sub folder
                                                        newPath = currentPath + "/" + newPath;
                                                    }

                                                }

                                            }
                                            if (newFileName != currentFileName || newPath != currentPath)
                                            {
                                                #region rename or move file
                                                try
                                                {
                                                    // move the file
                                                    item.File.MoveTo(newPath + "/" + newFileName, MoveOperations.Overwrite);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Moved '" + currentFileName + "' to '" + newFileName + "'");
                                                }
                                                catch (Exception ex)
                                                {
                                                    log.Error("Error moving or renaming file", ex);
                                                }


                                                #endregion
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


                            log.Info("Finished task '" + taskOptions.Title + "'");

                        }
                    
                        loopCnt--;
                        if (loopCnt > 0)
                        {
                            // delay next run
                            if (runOptions.LoopDelay > 0)
                            {
                                log.Info("Delay next run for " + runOptions.LoopDelay + "sec");
                                Thread.Sleep(runOptions.LoopDelay * 1000);
                            }
                        }
                    }
 





                }
            }
            catch (Exception ex)
            {
                log.Error("Exception : " + ex.Message);
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

        private static string GetFieldValueAsString(ListItem item, string fieldName)
        {
            string erg = "";
            object intOrString = GetFieldValueAsIntOrString(item, fieldName);
            if (intOrString != null)
            {
                erg = intOrString.ToString();
            }
            return erg;
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
