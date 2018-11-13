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

            RunOptions options;
            try
            {
                options = RunOptions.LoadFromXMl(xmlFileName);
            }
            catch (Exception ex)
            {
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
                log.Error("Missing SiteUrl in xmlFile: " + xmlFileName);
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



                    // loading all field definitions from sharepoint
                    FieldCollection fields = ctx.Web.Fields;
                    ctx.Load(fields);

                    // try to get all static uptade field values
                    List<FieldInfo> updateFields = new List<FieldInfo>();
                    foreach (var field in options.UpdateFields)
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
                                    return;
                                }
                                var statusTermId = TaxonomyHelpers.GetTermIdForTerm(field.Format, txField.TermSetId, ctx);
                                log.Info("field to update is a taxonomy field! id found:" + statusTermId);
                                if (statusTermId != null)
                                {
                                    field.NewValue = new TaxonomyFieldValue() {
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
                                return;
                            }
                        }
                        
                    }

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
                        workflowRep = new WorkflowRepository(ctx, spList.Id);
                    }

                    // build viewFields
                    string viewFields;
                    viewFields = "<ViewFields>";
       
                    foreach(var field in options.QueryFields)
                    {
                        viewFields += "<FieldRef " +
                            "Name='" + field.FieldName + "'" +
                            " />";
                    }
                    foreach (var field in options.UpdateFields)
                    {
                        viewFields += "<FieldRef " +
                            "Name='" + field.FieldName + "'" +
                            " />";
                    }

                    viewFields += "</ViewFields>";

                    string scope = (!string.IsNullOrEmpty(options.Scope)) ? " Scope='" + options.Scope + "' " : "";
 
                    // build caml query
                    CamlQuery camlQuery = new CamlQuery();
                    
                    camlQuery.ViewXml = "<View "+ scope + " > " +
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
                            foreach (var field in options.QueryFields)
                            {
                                object fieldValue = GetFieldValueAsIntOrString(item, field.FieldName); 
                                if (field.FieldName == "_UIVersionString")
                                {
                                    // prepare the version value as a future major version
                                    // because it is impossible to update a major version!
                                    decimal dec;
                                    if (decimal.TryParse(fieldValue.ToString().Replace(".",","), out dec)) {
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
                            string currentFileName = null;
                            string newFileName = null;
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
                                            foreach(var field in updateFields)
                                            {
                                                if (field.CurrentValueAsString != field.NewValueAsString)
                                                {
                                                    if(field.SpField.TypeAsString == "TaxonomyFieldType")
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
                                                    realoadedItem.File.CheckIn(options.CheckinMessage, options.CheckinType);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Checked in: '" + newFileName + "'");
                                                }
                                                #endregion
                                                #region publish
                                                if (!string.IsNullOrEmpty(options.PublishInfo))
                                                {
                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    realoadedItem.File.Publish(options.PublishInfo);
                                                    ctx.ExecuteQuery();
                                                    log.Info("Published: '" + newFileName + "'");
                                                }
                                                #endregion
                                                #region approve
                                                if (!string.IsNullOrEmpty(options.ApproveInfo))
                                                {
                                                    var realoadedItem = spList.GetItemById(item.Id);
                                                    realoadedItem.File.Approve(options.ApproveInfo);
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

                                    // find current and new file name
                                    currentFileName = item["FileLeafRef"].ToString();
                                    newFileName = currentFileName;
                                    if (!string.IsNullOrEmpty(options.FileNameFormat))
                                    {
                                        string extension = System.IO.Path.GetExtension(currentFileName);
                                        newFileName = string.Format(options.FileNameFormat, fieldValues.ToArray()) + extension;
                                        // replace illegal characters
                                        newFileName = string.Join("_", newFileName.Split(System.IO.Path.GetInvalidFileNameChars()));
                                    }

                                    string currentPath = "";
                                    string newPath = "";
                                    if (newFileName != currentFileName || !string.IsNullOrEmpty(options.MoveTo))
                                    {
                                        // get current and new path
                                        
                                        currentPath = item["FileRef"].ToString().Replace("/" + currentFileName, "");
                                        newPath = currentPath;
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

                                    }
                                    if (newFileName != currentFileName || newPath != currentPath)
                                    {
                                        #region rename or move file
                                        try
                                       
                                        {
                                            // move the file
                                            item.File.MoveTo(newPath + "/" + newFileName, MoveOperations.Overwrite);
                                            ctx.ExecuteQuery();
                                            log.Info("Moved ''" + currentFileName + "' to '" + newFileName + "'");
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
