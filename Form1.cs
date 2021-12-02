using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Word = Microsoft.Office.Interop.Word;
namespace WordToPDFConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Worked but generate watermarks
        //private void button1_Click(object sender, EventArgs e)
        //{

        //    //MessageBox.Show("to be implemented");
        //    using (var fbd = new FolderBrowserDialog())
        //    {
        //        DialogResult result = fbd.ShowDialog();

        //        if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
        //        {

        //            var files = Directory.GetFiles(fbd.SelectedPath, "*.docx").ToList();

        //            foreach (string filePath in files)

        //            {

        //                //Get File Name
        //                List<string> splittedpath = filePath.Split('\\').ToList();
        //                string fileName = splittedpath.Last().Split('.')[0];
        //                //Load Document
        //                Document document = new Document();
        //                document.LoadFromFile(filePath);
        //                //exclude file name form path 
        //                string newfilePath = string.Join("\\", splittedpath.Take(splittedpath.Count - 1));
        //                //Convert Word to PDF
        //                document.SaveToFile($"{newfilePath}/{fileName}.PDF", FileFormat.PDF);

        //                //Launch Document
        //                System.Diagnostics.Process.Start($"{newfilePath}/{fileName}.PDF");
        //            }
        //        }
        //    }

        //}
        #endregion
        private void ConvertAllDocToPDF_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("to be implemented");
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {

                    var files = Directory.GetFiles(fbd.SelectedPath, "*.docx").ToList();

                    foreach (string filePath in files)
                    {

                        convertDOCtoPDF(filePath);
                    }

                    MessageBox.Show("Done");
                }
            }

        }

        /// <summary>
        /// the Original source of this function is https://stackoverflow.com/a/55334441.
        /// </summary>
        /// <param name="filePath"></param>
        private void convertDOCtoPDF(string filePath)
        {
            //Get File Name
            List<string> splittedpath = filePath.Split('\\').ToList();
            string fileName = splittedpath.Last().Split('.')[0];

            //exclude file name form path 
            string newfilePath = string.Join("\\", splittedpath.Take(splittedpath.Count - 1));

            //String PATH_APP_PDF = newfilePath;

            object misValue = System.Reflection.Missing.Value;
            var WORD = new Word.Application();

            Word.Document doc = WORD.Documents.Open(filePath);
            doc.Activate();

            doc.SaveAs2($"{ newfilePath}/{ fileName}.PDF", Word.WdSaveFormat.wdFormatPDF, misValue, misValue, misValue,
            misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            doc.Close();
            WORD.Quit();


            releaseObject(doc);
            releaseObject(WORD);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                //TODO
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
