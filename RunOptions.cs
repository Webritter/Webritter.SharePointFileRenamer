using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Webritter.SharePointFileRenamer
{
   
    public class RunOptions
    {

        public int Id { get; set; }
        public bool Enabled { get; set; }
        public string Domain { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SiteUrl { get; set; }
        public string LibraryName { get; set; }
        public string CamlQuery { get; set; }
        public string FileNameFormat { get; set; }
        public List<FieldOptions> FieldNames { get; set; }

        public string StatusFieldName { get; set; }
        public string StatusSuccessValue { get; set; }
        public string CheckinMessage { get; set; }
        public CheckinType CheckinType { get; set; }
        public string PublishInfo { get; set; }
        public string ApproveInfo { get; set; }

        // constuctor
        public RunOptions()
        {
            FieldNames =  new List<FieldOptions>();
        }

        public static RunOptions LoadFromXMl(string xmlFileName)
        {
            // Now we can read the serialized book ...  
            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(RunOptions));
            System.IO.StreamReader file = new System.IO.StreamReader(xmlFileName);
            RunOptions result = (RunOptions)reader.Deserialize(file);
            file.Close();
            return result;
        }

        // file load and save
        public void SaveAsXml(string xmlFileName)
        {
              
            var writer = new System.Xml.Serialization.XmlSerializer(typeof(RunOptions));
            var wfile = new System.IO.StreamWriter(xmlFileName);
            writer.Serialize(wfile, this);
            wfile.Close();
        }

        public static void GreateSampleXml(string filename)
        {
            RunOptions sample = new RunOptions()
            {
                SiteUrl = "https://sharepoint.identecsolutions.com/sites/WS0029",
                Domain = "",
                Username = "identec\\communardo",
                Password = "secret",
                LibraryName = "Product Customization",
                FileNameFormat = "PCF_{0:00000}",
                CamlQuery = @"
                            <Where>
                                  <And>
                                     <IsNull>
                                        <FieldRef Name='identecDocumentStatus' />
                                     </IsNull>
                                     <Eq>
                                        <FieldRef Name='FSObjType' />
                                        <Value Type='Integer'>0</Value>
                                     </Eq>
                                  </And>
                               </Where>",
//                CamlQuery = "<Where><And><Eq><FieldRef Name='FileDirRef' /><Value Type='Text'>/sites/WS0029/Product Customization</Value></Eq><IsNull><FieldRef Name='identecDocumentStatus'/></IsNull></And></Where>",
                FieldNames = new List<FieldOptions>()
                {
                    new FieldOptions()
                    {
                        FieldName = "ID",
                        ShouldNotBeNull = true
                    }
                },

                StatusFieldName = "identecDocumentStatus",
                StatusSuccessValue = "Draft",
                CheckinMessage ="File Renamed",
                CheckinType = CheckinType.OverwriteCheckIn,
                PublishInfo = null,
                ApproveInfo = null

            };
            sample.SaveAsXml(filename);
        }
    }

    public class FieldOptions
    {
        public string FieldName { get; set; }
        public bool ShouldNotBeNull { get; set; }
    }
}
