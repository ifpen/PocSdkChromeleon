using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Transactions;
using Thermo.Chromeleon.Sdk.Interfaces.Data;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments.Queue;

namespace PocSdkChromeleon
{   
    internal class SequenceRun
    {
        private IItemFactory ItemFactory { get;}
        private ISequence Sequence { get; set; }
        private IInstrument Instrument { get; set; }
        private SynchronizationContext SynchronizationContext { get;}

        public SequenceRun(IItemFactory itemFactory, SynchronizationContext synchronizationContext)
        {
            ItemFactory = itemFactory;
            SynchronizationContext = synchronizationContext;
        }

        internal void CreateSequence(String destinationFolder)
        {
            using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.MaxValue))
            {
                // Emplacement de la séquence à créer dans l'arborescence Chromeleon. Probablement à paramétrer.
                var folderUri = new Uri(destinationFolder);

                if (ItemFactory.TryGetItem(folderUri, out IFolder folder))
                {
                    // Create the sequence and add it to the data vault
                    var sequence = ItemFactory.CreateSequence("testSDK" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), folder);
                    // Create new injections and add them to the sequence
                    var firstInjection = ItemFactory.CreateInjection("first injection", sequence);
                   // var secondInjection = ItemFactory.CreateInjection("second injection", sequence);

                    // Load a processing method and an instrument method
                    IProcessingMethod procMeth;
                    var procMethUrl = new Uri("chrom://isntsv-pacha4/DATA_R06_sql/DSI/TEST_PROCESSING_METHOD.procmeth");
                    if (!ItemFactory.TryGetItem(procMethUrl, out procMeth))
                        throw new Exception("processing method not found");

                    IInstrumentMethod instMeth;
                    var instMethUrl = new Uri("chrom://isntsv-pacha4/DATA_R06_sql/DSI/TEST_INSTRUMENT_METHOD.instmeth");
                    if (!ItemFactory.TryGetItem(instMethUrl, out instMeth))
                        throw new Exception("instrument method not found");

                    IReportTemplate reportTemplate;
                    var reportTemplateUrl = new Uri("chrom://isntsv-pacha4/DATA_R06_sql/DSI/TEST_REPORT_TEMPLATE.report");
                    if (!ItemFactory.TryGetItem(reportTemplateUrl, out reportTemplate))
                        throw new Exception("report template not found");

                    // Copy the methods to the sequence
                    procMeth.CopyTo(sequence, CopyOptions.CurrentVersion);
                    instMeth.CopyTo(sequence, CopyOptions.CurrentVersion);
                    reportTemplate.CopyTo(sequence, CopyOptions.CurrentVersion);
                    sequence.DefaultReportTemplateName = reportTemplate.Name;

                    // Assing the methods to the injections and apply a new inject volume
                    foreach (var injection in sequence.Injections)
                    {
                        injection.ProcessingMethodName.Value = procMeth.Name;
                        injection.InstrumentMethodName.Value = instMeth.Name;
                        injection.InjectionVolume.Value = 30;
                    }
                    // Complete the transaction and store the new sequence
                    scope.Complete();
                    Sequence = sequence;
                }
            }
        }

        internal void AddToInstrument(IInstrument instrument)
        {
            Instrument = instrument;
            var connectTask = instrument.ConnectAsync();
            connectTask.Wait();
            if (instrument.Connection == Progress.Completed)
            {
                instrument.QueueControl.AddItem(Sequence.Url);
            }
        }
        internal void SetExportAfterInjection()
        {
            Instrument.QueueControl.InjectionRunEnded += new EventHandler<InjectionChangedEventArgs>(OnInjectionEnded);
        }

        private void OnInjectionEnded(object sender, InjectionChangedEventArgs args) {
            SynchronizationContext.Post(_ =>
            {
                //Nécéssaire, car l'objet dans la mémoire n'a pas les données de l'injection qui vient de se terminer.
                Sequence.Reload();
                var injections = Sequence.Injections.Where(i => i.Url.Segments.Last() == args.InjectionUri.Segments.Last());
                var channel = Sequence.DefaultChannel == String.Empty ? injections.FirstOrDefault().GetAvailableChannelNames(false, false).FirstOrDefault() : Sequence.DefaultChannel;
                Sequence.DefaultReportTemplate.ExportToAndi(injections, channel, @"C:\Export\{seq.name}\{injection.name}_{injection.number;""00""}", null);
            }, null);
        }

        internal void StartInstrument()
        {

            var check = Instrument.QueueControl.ForceQueueStart();
        }
    }
}
