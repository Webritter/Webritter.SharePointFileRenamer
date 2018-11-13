using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Webritter.SharePointFileRenamer
{
    public class WorkflowRepository
    {
        private InteropService _interopService;
        private WorkflowInstanceService _instanceService;
        private WorkflowSubscriptionCollection _subscriptions;
        public ClientContext Context;

        public WorkflowRepository(ClientContext clientContext, Guid listId)
        {
            var Context = clientContext;

            var wfMngr = new WorkflowServicesManager(Context, Context.Web);
            var subsrService = wfMngr.GetWorkflowSubscriptionService();
            _interopService = wfMngr.GetWorkflowInteropService();
            _instanceService = wfMngr.GetWorkflowInstanceService();
            _subscriptions = subsrService.EnumerateSubscriptionsByList(listId);
            
            Context.Load(_subscriptions);
            Context.ExecuteQuery();
        }

        #region basic methods 
        public Guid RunSiteWorkflow45(string wfName, IDictionary<string, object> payload)
        {
            var subscription = _subscriptions.FirstOrDefault(sub => sub.Name == wfName);
            var wf = _instanceService.StartWorkflow(subscription, payload);

            Context.ExecuteQuery();
            return wf.Value;
        }

        public Guid RunListWorkflow45(int itemId, string wfName, IDictionary<string, object> payload)
        {
            var subscription = _subscriptions.FirstOrDefault(sub => sub.Name == wfName);
            if (payload == null) //StartWorkflowOnListItem will throw exception if null
                payload = new Dictionary<string, object>();

            var wf = _instanceService.StartWorkflowOnListItem(subscription, itemId, payload);

            Context.ExecuteQuery();
            return wf.Value;
        }
        #endregion
    }
}
