using System;
using System.Linq;
using System.Threading;
using System.Transactions;
using System.Windows.Forms;
using Thermo.Chromeleon.Sdk.Common;
using Thermo.Chromeleon.Sdk.Interfaces.Data;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments.Queue;

namespace PocSdkChromeleon
{
    internal class Program
    {
        static void Main(string[] args)
        {
            SynchronizationContext context = SynchronizationContext.Current;
            CmSdkScope sdkScope = new CmSdkScope();
            var roles = CmSdk.Logon.GetRoles("thermo");
            CmSdk.Logon.DoSilentLogon("thermo", "thermo", roles.First());
            var itemFactory = CmSdk.GetItemFactory();
            var uiFactory = CmSdk.GetUserInterfaceFactory();

            SequenceRun sequenceRun = new SequenceRun(itemFactory, context);

            sequenceRun.CreateSequence("chrom://isntsv-pacha4/DATA_R06_sql/DSI/");

            var instrumentSelectorDialog = uiFactory.CreateInstrumentSelectorDialog();
            if (instrumentSelectorDialog.ShowDialog() == DialogResult.OK)
            {
                IInstrument instrument = instrumentSelectorDialog.SelectedInstrument;
                sequenceRun.AddToInstrument(instrument);
                sequenceRun.SetExportAfterInjection();
                sequenceRun.StartInstrument();
            }
        }
    }
}
