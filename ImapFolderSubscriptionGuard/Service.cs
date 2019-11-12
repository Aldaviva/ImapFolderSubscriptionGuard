using System.ServiceProcess;

namespace ImapFolderSubscriptionGuard {

    internal partial class Service: ServiceBase {

        private SubscriptionGuard subscriptionGuard;

        public Service() {
            InitializeComponent();
        }

        protected override void OnStart(string[] args) {
            subscriptionGuard = new SubscriptionGuard();
            subscriptionGuard.startGuarding();
        }

        protected override void OnStop() {
            subscriptionGuard.stopGuarding();
        }

    }

}