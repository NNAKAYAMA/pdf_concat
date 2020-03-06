using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
namespace PDFConcat
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<PdfFile> joinPDFs;
        public MainWindow()
        {
            InitializeComponent();
            joinPDFs = new ObservableCollection<PdfFile>();
            listView_pdf_files.ItemsSource =  joinPDFs;
        }

        
        public class PdfFile
        {
            public string Path
            {
                get;
                set;
            }
            public string Title
            {
                get;
                set;
            }
            public int id
            {
                get;
                set;
            }
        }

        public void PDF_Drop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null)
            {
                return;
            }
            for (int i = 0; i < files.Length; i++)
            {
                if (System.IO.File.Exists(files[i]) == false)
                {
                    continue;
                }
                if (System.IO.Path.GetExtension(files[i]) != ".pdf"){
                    continue;
                }
                PdfFile pdf = new PdfFile();
                pdf.Title = System.IO.Path.GetFileName(files[i]);
                pdf.Path = System.IO.Path.GetFullPath(files[i]);
                pdf.id = joinPDFs.Count + 1;
                joinPDFs.Add(pdf);
            }

        }

        private void Button_Increase_Item_Order(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if (selectedID == (joinPDFs.Count - 1) || selectedID == -1)
            {
                return;
            }
            joinPDFs[selectedID].id += 1;
            joinPDFs[selectedID + 1].id -= 1;
            joinPDFs = new ObservableCollection<PdfFile>(joinPDFs.OrderBy(n => n.id));
            listView_pdf_files.ItemsSource = joinPDFs;

        }

        private void Button_Decrease_Item_Order(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if(selectedID == 0 || selectedID == -1)
            {
                return;
            }
            joinPDFs[selectedID].id -= 1;
            joinPDFs[selectedID - 1].id += 1;
            joinPDFs = new ObservableCollection<PdfFile>(joinPDFs.OrderBy(n => n.id));
            listView_pdf_files.ItemsSource = joinPDFs;
        }

        private void Button_Remove_Item(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if (selectedID != -1)
            {
                joinPDFs.RemoveAt(selectedID);
            }
        }

        private void Button_Concat_PDF_Click(object sender, RoutedEventArgs e)
        {
            if(joinPDFs.Count == 0)
            {
                return;
            }
            SaveFileDialog sdialog = new SaveFileDialog();
            sdialog.Title = "保存する場所を指定してください";
            sdialog.DefaultExt = "pdf";
            sdialog.FileName = DateTime.Now.ToString("yymmddHHMMss") + ".pdf";
            Nullable<bool> result = sdialog.ShowDialog();
            if(result == false)
            {
                return;
            }
            FileStream joinStream = new  FileStream(sdialog.FileName, FileMode.OpenOrCreate);
            Document joinDocument = new Document();
            PdfWriter joinWriter = PdfWriter.GetInstance(joinDocument, joinStream);
            joinDocument.Open();
            PdfContentByte joinPcb = joinWriter.DirectContent;
            foreach(PdfFile pdf in joinPDFs)
            {
                PdfReader pdfReader = new PdfReader(File.ReadAllBytes(pdf.Path));
                int joinNp = pdfReader.NumberOfPages;
                for(int joinPageNum = 1;joinPageNum <= joinNp; joinPageNum++)
                {
                    int pageRotation = pdfReader.GetPageRotation(1);
                    joinDocument.SetPageSize(pdfReader.GetPageSizeWithRotation(joinPageNum));
                    joinDocument.NewPage();
                    PdfImportedPage joinPage = joinWriter.GetImportedPage(pdfReader, joinPageNum);
                    if(pageRotation == 90)
                    {
                        joinPcb.AddTemplate(joinPage, 0, -1, 1, 0, 0, pdfReader.GetPageSizeWithRotation(joinPageNum).Height);
                    }
                    else if (pageRotation == 180)
                    {
                        joinPcb.AddTemplate(joinPage, -1, 0,1, -1, pdfReader.GetPageSizeWithRotation(joinPageNum).Width, pdfReader.GetPageSizeWithRotation(joinPageNum).Height);
                    }
                    else if (pageRotation == 270)
                    {
                        joinPcb.AddTemplate(joinPage, 0, 1, -1, 0, pdfReader.GetPageSizeWithRotation(joinPageNum).Width,0);
                    }
                    else
                    {
                        joinPcb.AddTemplate(joinPage, 1, 0, 0,1, 0, 0);
                    }
                }
            }
            joinDocument.Close();
            joinPDFs.Clear();
        }

        private void button_clear_list_Click(object sender, RoutedEventArgs e)
        {
            joinPDFs.Clear();
        }
    }
}
