using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using Org.BouncyCastle.Bcpg;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;

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
            listView_pdf_files.ItemsSource = joinPDFs;
        }

        private string SaveFileName;
        private string SaveDirectoryPath;

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

        public void _PDF_Drop(string[] files)
        {
            Array.Sort(files);
            if (files == null)
            {
                return;
            }
            for (int i = 0; i < files.Length; i++)
            {
                if (System.IO.Directory.Exists(files[i]))
                {
                    _PDF_Drop(Directory.GetFiles(files[i], "*", SearchOption.AllDirectories));
                    continue;
                }
                if (System.IO.Path.GetExtension(files[i]) != ".pdf")
                {
                    continue;
                }
                PdfFile pdf = new PdfFile();
                pdf.Title = System.IO.Path.GetFileName(files[i]);
                pdf.Path = System.IO.Path.GetFullPath(files[i]);
                pdf.id = joinPDFs.Count + 1;
                joinPDFs.Add(pdf);
            }

        }

        public void PDF_Drop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files.Length == 0) return;
            if (files.Length == 1 && Directory.Exists(files[0]))
            {
                SaveFileName = new DirectoryInfo(files[0]).Name;
            }
            SaveDirectoryPath = Path.GetDirectoryName(files[0]);
            _PDF_Drop(e.Data.GetData(DataFormats.FileDrop) as string[]);

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
            if (selectedID == 0 || selectedID == -1)
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
            if (joinPDFs.Count == 0)
            {
                return;
            }
            Concat_PDF();
            joinPDFs.Clear();
        }

        private void button_clear_list_Click(object sender, RoutedEventArgs e)
        {
            joinPDFs.Clear();
        }

        private void Concat_PDF()
        {
            SaveFileDialog sdialog = new SaveFileDialog();
            sdialog.InitialDirectory = SaveDirectoryPath;
            sdialog.Title = "保存する場所を指定してください";
            sdialog.DefaultExt = "pdf";
            sdialog.FileName = SaveFileName ?? DateTime.Now.ToString("yymmddHHMMss") + ".pdf";
            var result = sdialog.ShowDialog();
            if (result == false)
            {
                return;
            }
            PdfReader.unethicalreading = true;
            FileStream joinStream = new FileStream(sdialog.FileName, FileMode.OpenOrCreate);
            Document joinDocument = new Document();
            PdfWriter joinWriter = PdfWriter.GetInstance(joinDocument, joinStream);
            joinDocument.Open();
            PdfContentByte joinPcb = joinWriter.DirectContent;
            var fontName = "meiryo";
            if (!FontFactory.IsRegistered(fontName))
            {
                var fontPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\meiryo.ttc";
                FontFactory.Register(fontPath, fontName);
            }
            Font pfont =
                    FontFactory.GetFont(fontName,
                    BaseFont.IDENTITY_H,    //横書き
                    BaseFont.EMBEDDED,  //フォントをPDFファイルに組み込まない（重要）
                    10f,                    //フォントサイズ
                    Font.NORMAL,           //フォントスタイル
                    BaseColor.BLACK);       //フォントカラー
            Font h1font =
              FontFactory.GetFont(fontName,
              BaseFont.IDENTITY_H,    //横書き
              BaseFont.EMBEDDED,  //フォントをPDFファイルに組み込まない（重要）
              14f,                    //フォントサイズ
              Font.NORMAL,           //フォントスタイル
              BaseColor.BLACK);
            int pageCount = 2;

            joinDocument.NewPage();
            string parentDirNameBuff = new DirectoryInfo(joinPDFs[0].Path).Parent.Name;
            if (true)
            {
                joinDocument.Add(new Paragraph(parentDirNameBuff, h1font));
            }


            // 目次の作成
            foreach (PdfFile pdf in joinPDFs)
            {
                PdfReader pdfReader = new PdfReader(File.ReadAllBytes(pdf.Path));
                string parentDirName = new DirectoryInfo(pdf.Path).Parent.Name;

                if(parentDirName != parentDirNameBuff)
                {
                    joinDocument.Add(new Paragraph(20f, parentDirName, h1font));
                    parentDirNameBuff = parentDirName;
                }
                Chunk link = new Chunk(pdf.Title, pfont);
                PdfAction action = PdfAction.GotoLocalPage(pageCount, new PdfDestination(PdfDestination.FIT), joinWriter);
                link.SetAction(action);
                joinDocument.Add(new Paragraph(link));
                pageCount += pdfReader.NumberOfPages;
            }

            foreach (PdfFile pdf in joinPDFs)
            {
                
                PdfReader pdfReader = new PdfReader(File.ReadAllBytes(pdf.Path));

                int joinNp = pdfReader.NumberOfPages;
                for (int joinPageNum = 1; joinPageNum <= joinNp; joinPageNum++)
                {
                    int pageRotation = pdfReader.GetPageRotation(1);
                    joinDocument.SetPageSize(pdfReader.GetPageSizeWithRotation(joinPageNum));
                    joinDocument.NewPage();
                    PdfImportedPage joinPage = joinWriter.GetImportedPage(pdfReader, joinPageNum);
                    if (pageRotation == 90)
                    {
                        joinPcb.AddTemplate(joinPage, 0, -1, 1, 0, 0, pdfReader.GetPageSizeWithRotation(joinPageNum).Height);
                    }
                    else if (pageRotation == 180)
                    {
                        joinPcb.AddTemplate(joinPage, -1, 0, 1, -1, pdfReader.GetPageSizeWithRotation(joinPageNum).Width, pdfReader.GetPageSizeWithRotation(joinPageNum).Height);
                    }
                    else if (pageRotation == 270)
                    {
                        joinPcb.AddTemplate(joinPage, 0, 1, -1, 0, pdfReader.GetPageSizeWithRotation(joinPageNum).Width, 0);
                    }
                    else
                    {
                        joinPcb.AddTemplate(joinPage, 1, 0, 0, 1, 0, 0);
                    }
                }
            }
            joinDocument.Close();
            System.Diagnostics.Process.Start(sdialog.FileName);
        }

        
    }
}
