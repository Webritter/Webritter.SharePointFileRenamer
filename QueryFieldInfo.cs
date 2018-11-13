using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Webritter.SharePointFileRenamer
{
    public class FieldInfo
    {
        public Field SpField { get; set; }
        public string CurrentValueAsString { get; set; }
        public TaxonomyFieldValue NewValue { get; set; }
        public string NewValueAsString { get; set; }
        public string Format { get; set; }
        public bool IsStaticValue { get; set; }


    }
}
