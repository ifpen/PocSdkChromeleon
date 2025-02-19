using System;
using System.Linq;
using System.Threading;
using System.Transactions;
using System.Windows.Forms;
using Thermo.Chromeleon.Sdk.Common;
using Thermo.Chromeleon.Sdk.Interfaces.Data;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments;
using Thermo.Chromeleon.Sdk.Interfaces.Instruments.Queue;
using Thermo.Chromeleon.Sdk.Interfaces.UserInterface;

namespace AppPocChromeleon
{
    public partial class Run : Form
    {
        private IUserInterfaceFactory userInterfaceFactory;
        private IItemFactory itemFactory;
        private IInstrumentAccess instrumentAccess;
        private ISequence sequence;
        private SynchronizationContext synchronizationContext;
        public Run()
        {
            InitializeComponent();
        }

        private void Run_Load(object sender, EventArgs e)
        {
            CmSdk.Logon.DoLogon();
            userInterfaceFactory = CmSdk.GetUserInterfaceFactory();
            itemFactory = CmSdk.GetItemFactory();
            instrumentAccess = CmSdk.GetInstrumentAccess();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void ParcourirSequenceTextBox_Click(object sender, EventArgs e)
        {
            var sequencePicker = userInterfaceFactory.CreateObjectPickerDialog();
            sequencePicker.AddFilter(typeof(IFolder));
            sequencePicker.InitialUrl = GetUriFromTextOrNull(sequenceFolderTextBox.Text);
            if (sequencePicker.ShowDialog() == DialogResult.OK)
            {
                sequenceFolderTextBox.Text = sequencePicker.SelectedParentUrl.AbsoluteUri;
            }

        }

        private void ParcourirInstrumentMethodTextBox_Click(object sender, EventArgs e)
        {
            var instrumentMethodPicker = userInterfaceFactory.CreateObjectPickerDialog();
            instrumentMethodPicker.AddFilter(typeof(IInstrumentMethod));
            instrumentMethodPicker.InitialUrl = GetUriFromTextOrNull(InstrumentMethodTextBox.Text);
            if (instrumentMethodPicker.ShowDialog() == DialogResult.OK)
            {
                InstrumentMethodTextBox.Text = instrumentMethodPicker.SelectedUrls.FirstOrDefault()?.AbsoluteUri;
            }
        }

        private Uri GetUriFromTextOrNull(String text)
        {
            if (Uri.TryCreate(text, UriKind.Absolute, out var uri))
            {
                return uri;
            }
            return null;
        }

        private void ParcourirProcessingMethodTextBox_Click(object sender, EventArgs e)
        {
            var processingMethodPicker = userInterfaceFactory.CreateObjectPickerDialog();
            processingMethodPicker.AddFilter(typeof(IProcessingMethod));
            processingMethodPicker.InitialUrl = GetUriFromTextOrNull(ProcessingMethodTextBox.Text);
            if (processingMethodPicker.ShowDialog() == DialogResult.OK)
            {
                ProcessingMethodTextBox.Text = processingMethodPicker.SelectedUrls.FirstOrDefault()?.AbsoluteUri;
            }
        }

        private void ParcourirReportingTemplateTextBox_Click(object sender, EventArgs e)
        {
            var reportingTemplatePicker = userInterfaceFactory.CreateObjectPickerDialog();
            reportingTemplatePicker.AddFilter(typeof(IReportTemplate));
            reportingTemplatePicker.InitialUrl = GetUriFromTextOrNull(ReportingTemplateTextBox.Text);
            if (reportingTemplatePicker.ShowDialog() == DialogResult.OK)
            {
                ReportingTemplateTextBox.Text = reportingTemplatePicker.SelectedUrls.FirstOrDefault()?.AbsoluteUri;
            }

        }

        private void ParcourirInstrumentTextBox_Click(object sender, EventArgs e)
        {
            var instrumentSelectorDialog = userInterfaceFactory.CreateInstrumentSelectorDialog();
            if (instrumentSelectorDialog.ShowDialog() == DialogResult.OK)
            {
                IInstrument instrument = instrumentSelectorDialog.SelectedInstrument;
                InstrumentTextBox.Text = instrument.ApplicationUri.AbsoluteUri;
            }
        }

        private void LancerSequenceButton_Click(object sender, EventArgs e)
        {
            synchronizationContext = SynchronizationContext.Current;
            CreateSequence();
            AddToInstrumentAndRun();

        }

        private void CreateSequence()
        {
            using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.MaxValue))
            {
                // Emplacement de la séquence à créer dans l'arborescence Chromeleon. Probablement à paramétrer.
                var folderUri = new Uri(sequenceFolderTextBox.Text);

                if (itemFactory.TryGetItem(folderUri, out IFolder folder))
                {
                    // Create the sequence and add it to the data vault
                    sequence = itemFactory.CreateSequence("testSDK" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), folder);
                    // Create new injections and add them to the sequence
                    var firstInjection = itemFactory.CreateInjection("first injection", sequence);
                    // var secondInjection = ItemFactory.CreateInjection("second injection", sequence);

                    // Load a processing method and an instrument method
                    IProcessingMethod procMeth;
                    var procMethUrl = new Uri(ProcessingMethodTextBox.Text);
                    if (!itemFactory.TryGetItem(procMethUrl, out procMeth))
                        throw new Exception("processing method not found");

                    IInstrumentMethod instMeth;
                    var instMethUrl = new Uri(InstrumentMethodTextBox.Text);
                    if (!itemFactory.TryGetItem(instMethUrl, out instMeth))
                        throw new Exception("instrument method not found");

                    IReportTemplate reportTemplate;
                    var reportTemplateUrl = new Uri(ReportingTemplateTextBox.Text);
                    if (!itemFactory.TryGetItem(reportTemplateUrl, out reportTemplate))
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
                }
            }
        }
        private void AddToInstrumentAndRun()
        {
            if (!instrumentAccess.TryFindInstrument(new Uri(InstrumentTextBox.Text), out IInstrument instrument))
                throw new Exception("instrument not found");


            instrument.TakeControl();
            instrument.QueueControl.AddItem(sequence.Url);

            instrument.QueueControl.InjectionRunEnded += new EventHandler<InjectionChangedEventArgs>(OnInjectionEnded);
            var check = instrument.QueueControl.ForceQueueStart();
        }

        private void OnInjectionEnded(object sender, InjectionChangedEventArgs args)
        {
            synchronizationContext.Post(_ =>
            {
                //Nécéssaire, car l'objet dans la mémoire n'a pas les données de l'injection qui vient de se terminer.
                sequence.Reload();
                var injections = sequence.Injections.Where(i => i.Url.Segments.Last() == args.InjectionUri.Segments.Last());
                var channel = 
                sequence.DefaultChannel == String.Empty ? injections.FirstOrDefault().GetAvailableChannelNames(false, false).FirstOrDefault() : sequence.DefaultChannel;
                sequence.DefaultReportTemplate.ExportToAndi(injections, channel, @"C:\Export\{seq.name}\{injection.name}_{injection.number;""00""}", null);
            }, null);
        }
    }
}
