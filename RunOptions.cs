using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Webritter.SharePointFileRenamer
{
    public class RunOptions
    {
        // login informations
        public string Domain { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SiteUrl { get; set; }


        // tasks
        public List<RunOptionsTask> Tasks { get; set; }

        public RunOptions()
        {
            Tasks = new List<RunOptionsTask>();
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
            RunOptions sample = new RunOptions();

            sample.SiteUrl = "https://communardodemo.sharepoint.com/sites/HelmutsSpielwiese";

            // create renamer task    
            RunOptionsTask renamer = new RunOptionsTask()
            {
                Title = "Renamer",
                Enabled = true,
                LibraryName = "RenamerTest",
                FileNameFormat = "PCF_{0:00000}",
                CamlQuery = @"
		            <Where>
			            <And>
				            <Or>
					             <Eq>
						            <FieldRef Name='_ModerationStatus' />
						            <Value Type='ModStat'>2</Value>
					             </Eq>
					             <Eq>
						            <FieldRef Name='_ModerationStatus' />
						            <Value Type='ModStat'>3</Value>
					             </Eq>
 				            </Or>
				            <And>
					            <Eq>
						            <FieldRef Name='FSObjType' />
						            <Value Type='Integer'>0</Value>
					             </Eq>
					            <IsNull>
						            <FieldRef Name='CheckoutUser' />
					            </IsNull>
	 			            </And>
			            </And>
		            </Where>	
                ",
                QueryFields = new List<QueryFieldOptions>()
                {
                    new QueryFieldOptions()
                    {
                        FieldName = "ID",
                        ShouldNotBeNull = true
                    },
                     new QueryFieldOptions()
                    {
                        FieldName = "_UIVersionString"
                    }
                },
                UpdateFields = new List<UpdateFieldOptions>()
                {
                    new UpdateFieldOptions()
                    {
                        FieldName = "Title",
                        Format = "PCF_{0:00000}_{1}"
                    },
                    new UpdateFieldOptions()
                    {
                        FieldName = "DocumentVersion",
                        Format = "V{1}"
                    }

                },

                CheckinMessage = "File Renamed",
                CheckinType = CheckinType.OverwriteCheckIn,
                PublishInfo = null,
                ApproveInfo = null

            };

            sample.Tasks.Add(renamer);



            // create mover task
            RunOptionsTask mover = new RunOptionsTask()
            {
                Title = "Mover",
                Enabled = true,
                LibraryName = "RenamerTest",
                FileNameFormat = "PCF_{0:00000}",
                CamlQuery = @"
                    <Where>
                        <Eq>
                            <FieldRef Name='_ModerationStatus' />
                            <Value ype='ModStat' >Approved</Value>
                        </Eq>
                    </Where>
                ",
                QueryFields = new List<QueryFieldOptions>()
                {
                    new QueryFieldOptions()
                    {
                        FieldName = "ID",
                        ShouldNotBeNull = true
                    },
                     new QueryFieldOptions()
                    {
                        FieldName = "_UIVersionString"
                    }
                },
                MoveTo = "final"
            };

            sample.Tasks.Add(mover);

            sample.SaveAsXml(filename);
        }



    }
    public class RunOptionsTask
    {
        [XmlAttribute("Id")]
        public int Id { get; set; }
        [XmlAttribute("Name")]
        public string Title { get; set; }
        [XmlAttribute("Enabled")]
        public bool Enabled { get; set; }
        public string LibraryName { get; set; }
        public string CamlQuery { get; set; }
        public string Scope { get; set; }
        public string FileNameFormat { get; set; }
        public List<QueryFieldOptions> QueryFields { get; set; }
        public List<UpdateFieldOptions> UpdateFields { get; set; }
        public string MoveTo { get; set; }

        public string CheckinMessage { get; set; }
        public CheckinType CheckinType { get; set; }
        public string PublishInfo { get; set; }
        public string ApproveInfo { get; set; }

        // constuctor
        public RunOptionsTask()
        {
            QueryFields =  new List<QueryFieldOptions>();
            UpdateFields = new List<UpdateFieldOptions>();
        }


    }

    public class QueryFieldOptions
    {
        [XmlAttribute("Name")]
        public string FieldName { get; set; }
        [XmlAttribute("ShouldNotBeNull")]
        public bool ShouldNotBeNull { get; set; }
    }

    public class UpdateFieldOptions
    {
        [XmlAttribute("Name")]
        public string FieldName { get; set; }
        [XmlAttribute("Format")]
        public string Format { get; set; }
    }
}
