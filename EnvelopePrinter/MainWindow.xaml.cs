using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Windows.Xps.Packaging;
using System.IO;
using System.Printing;
using System.Windows.Xps;
using System.Globalization;
using System.Collections.ObjectModel;

namespace EnvelopePrinter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        MainViewModel MainViewModel { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            this.MainViewModel = new MainViewModel();
            this.DataContext = this.MainViewModel;

            this.MainViewModel.SearchText = "";


        }
                
        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            this.MainViewModel.SearchText = this.Search.Text;
        }

        private void PrintLabelButton_Click(object sender, RoutedEventArgs e)
        {

            FlowDocument flowDocument = new FlowDocument();
            //flowDocument.PageHeight = this.MainViewModel.EnvelopeHeightInPixels;
            //flowDocument.PageWidth = this.MainViewModel.EnvelopeWidthInPixels;

            flowDocument.Blocks.Add(new Paragraph() { FontSize=4.0 });

            Table byRpadTable = new Table() { Margin=new Thickness(20,0,0,0) };
            byRpadTable.Columns.Add(new TableColumn { Width = new GridLength( 240 ) });

            TableRowGroup trg = new TableRowGroup() { };

            TableRow tr = new TableRow();

            tr.Cells.Add(new TableCell(new Paragraph(new Run("By Registered Post with Ack. Due"))) { BorderThickness=new Thickness(0,0,0,1), BorderBrush=Brushes.Black });

            trg.Rows.Add(tr);

            byRpadTable.RowGroups.Add(trg);

            Table addressTable = new Table() { Margin = new Thickness(20, 0, 0, 0) };
            addressTable.Columns.Add(new TableColumn { Width = new GridLength(this.MainViewModel.FromSectionWidthInPixels) });
            addressTable.Columns.Add(new TableColumn { Width = new GridLength(this.MainViewModel.ToSectionWidthInPixels) });

            TableRowGroup trg1 = new TableRowGroup() { };

            TableRow tr1 = new TableRow();

            tr1.Cells.Add(new TableCell(new Paragraph(new Run("From")) { Margin = new Thickness(20, 120, 0, 0), FontWeight = FontWeights.Bold }) );

            TableCell cell2 = new TableCell(new Paragraph(new Run("To")) { Margin = new Thickness(50, 60, 0, 0), FontWeight = FontWeights.Bold }) { RowSpan = 2 };

            foreach (string toLine in this.MainViewModel.ToAddress)
                cell2.Blocks.Add(new Paragraph(new Run(toLine)) { Margin = new Thickness(70, 2, 0, 2) });


            tr1.Cells.Add(cell2);
            
            
            trg1.Rows.Add(tr1);

            TableRow tr2 = new TableRow();
            TableCell cell3 = new TableCell();

            foreach (string fromLine in this.MainViewModel.FromAddress)
                cell3.Blocks.Add(new Paragraph(new Run(fromLine)) { Margin = new Thickness(35, 0, 35, 0) });

            tr2.Cells.Add(cell3);

            trg1.Rows.Add(tr2);
            
            addressTable.RowGroups.Add(trg1);

            flowDocument.Blocks.Add(byRpadTable);
            flowDocument.Blocks.Add(addressTable);

            flowDocument.ColumnWidth = double.PositiveInfinity;

            TextRange dest = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);
            
            PrintDialog printDialog = new PrintDialog();

            printDialog.PrintTicket.PageOrientation = PageOrientation.Landscape;
            


            if (printDialog.ShowDialog() == true)
            {
                //Other settings
                DocumentPaginator paginator = ((IDocumentPaginatorSource)flowDocument).DocumentPaginator;
                //paginator.PageSize = new Size(); //Set the size of A4 here accordingly
                printDialog.PrintDocument(paginator, "My");

            }

            //DoThePrint(fd);

        }

        private double pixelsInOneMM = 3.77952755905512;
    }

    public class MainViewModel : INotifyPropertyChanged
    {
        public AddressViewModel [][] AddressForPrinting { get; set; }

        public List<AddressSelectionViewModel> AllAddresses { get; set; }
        private double pixelsInOneMM = 3.77952755905512;
        public List<AddressSelectionViewModel> AddressForSelection { get; set; }

        public ObservableCollection<string> FromAddress { get; set; }

        public ObservableCollection<string> ToAddress { get; set; }

        private double envelopeHeight;

        public double EnvelopeHeight
        {
            get { return envelopeHeight; }
            set { envelopeHeight = value; this.EnvelopeHeightInPixels = value * 10 * pixelsInOneMM; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("EnvelopeHeightInPixels")); }
        }

        private double envelopeWidth;

        public double EnvelopeWidth
        {
            get { return envelopeWidth; }
            set
            {
                envelopeWidth = value; 
                this.EnvelopeWidthInPixels = value * 10 * pixelsInOneMM; 
                this.FromSectionWidthInPixels = (3.0 / 8) * this.EnvelopeWidthInPixels; 
                this.ToSectionWidthInPixels = (5.0 / 8) * this.EnvelopeWidthInPixels;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("EnvelopeWidthInPixels"));
                    PropertyChanged(this, new PropertyChangedEventArgs("FromSectionWidthInPixels"));
                    PropertyChanged(this, new PropertyChangedEventArgs("ToSectionWidthInPixels"));
                }
            }
        }

        public double FromSectionWidthInPixels { get; set; }

        public double ToSectionWidthInPixels { get; set; }

        public double EnvelopeHeightInPixels { get; set; }

        public double EnvelopeWidthInPixels { get; set; }

        private string searchText = string.Empty;

        public string SearchText
        {
            get { return searchText; }
            set
            {
                searchText = value;

                if (!string.IsNullOrWhiteSpace(this.SearchText))
                {
                    this.AddressForSelection = this.AllAddresses.Where(a => a.Address.ToLower().Contains(this.SearchText.ToLower())).ToList();
                }
                else
                {
                    this.AddressForSelection = this.AllAddresses.ToList();
                }

                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("AddressForSelection"));
            }
        }

        public MainViewModel()
        {
            this.EnvelopeHeight = 10.2;

            this.EnvelopeWidth = 22.7;

            this.FromAddress = new ObservableCollection<string>
            {
                "Sender,",
                "Apt #, Some Apartments,",
                "Street Address,",
                "Area,",
                "City - 600001,",
                "State"
            };

            this.ToAddress = new ObservableCollection<string>
            {
                "Recipient",
                "Apt #, Some Apartments,",
                "Street Address,",
                "Area,",
                "City - 600001,",
                "State"
            };


            this.AddressForPrinting = new AddressViewModel[7][];

            for (int i = 0; i < 7; i++)
            {
                this.AddressForPrinting[i] = new AddressViewModel[3];
                for (int j = 0; j < 3; j++)
                    this.AddressForPrinting[i][j] = new AddressViewModel() { Address = string.Empty };
            }

            //XElement addresses = XElement.Load(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Addresses.xml"), LoadOptions.None);

            XElement addresses = new XElement("Addresses");

            this.AllAddresses = addresses.Descendants().Select
            (
                node=>
                new AddressSelectionViewModel
                {
                    Address=node.Value.Trim()
                }
            ).ToList();

        }


        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class ToCommandViewModel : ICommand
    {

        public AddressSelectionViewModel AddressSelectionViewModel { get; set; }

        public ToCommandViewModel(AddressSelectionViewModel asvm)
        {
            this.AddressSelectionViewModel = asvm;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            MainViewModel mainVM = parameter as MainViewModel;

            if (mainVM != null)
            {
                mainVM.ToAddress.Clear();
                this.AddressSelectionViewModel.Address.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList().ForEach(mainVM.ToAddress.Add);
            }

        }


        public event EventHandler CanExecuteChanged;
    }

    public class FromCommandViewModel : ICommand
    {

        public AddressSelectionViewModel AddressSelectionViewModel { get; set; }

        public FromCommandViewModel(AddressSelectionViewModel asvm)
        {
            this.AddressSelectionViewModel = asvm;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            MainViewModel mainVM = parameter as MainViewModel;

            if (mainVM != null)
            {
                mainVM.FromAddress.Clear();
                this.AddressSelectionViewModel.Address.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList().ForEach(mainVM.FromAddress.Add);
            }

        }


        public event EventHandler CanExecuteChanged;
    }

    public class AddressSelectionViewModel : INotifyPropertyChanged
    {

        public AddressSelectionViewModel()
        {
            this.AddCommand = new FromCommandViewModel(this);
            this.DeleteCommand = new ToCommandViewModel(this);
        }

        private string address;

        public string Address
        {
            get { return address; }
            set { address = value; }
        }

        public ICommand AddCommand { get; set; }
        public ICommand DeleteCommand { get; set; }


        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class AddressViewModel : INotifyPropertyChanged
    {
        private bool selected;

        public bool IsSelected
        {
            get { return selected; }
            set { selected = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsSelected")); }
        }


        private bool isEmpty;

        public bool IsEmpty
        {
            get { return string.IsNullOrWhiteSpace(this.Address); }
            set { isEmpty = value; if (string.IsNullOrWhiteSpace(this.Address)) this.CanSave = false; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsEmpty")); }
        }



        private Visibility Visible;

        public Visibility Visibility
        {
            get { return this.CanSave ? Visibility.Visible : Visibility.Hidden; }
            set {  }
        }

        private bool canSave;

        public bool CanSave
        {
            get { return canSave; }
            set { canSave = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Visibility")); }
        }
               

        private string address;

        public string Address
        {
            get { return address; }
            set { address = value; this.CanSave = !string.IsNullOrWhiteSpace(value); if (PropertyChanged != null) { PropertyChanged(this, new PropertyChangedEventArgs("Address")); PropertyChanged(this, new PropertyChangedEventArgs("IsEmpty")); } }
        }
        

        public event PropertyChangedEventHandler PropertyChanged;
    }

}
