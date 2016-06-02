using Microsoft.Office.Interop.Word;
using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace docxToPdfConverter
{
    public partial class MainPDFConverter : Form
    {
        public FilesToConvert fileList = new FilesToConvert();

        public MainPDFConverter()
        {
            InitializeComponent();
        }


        private void MainPDFConverter_Load(object sender, EventArgs e)
        {
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            int size = -1;
            DialogResult result = openFileDialog1.ShowDialog(); //Show the dialog.
            if (result == DialogResult.OK)    // Test result.
            {
                var files = openFileDialog1.FileNames;
                try
                {
                    textBoxDirPath.Text = "";
                    fileList = new FilesToConvert(files);

                    textBoxDirPath.Text = fileList.FilesToTextBox();
                }
                catch (IOException)
                {
                }
            }
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            var progress = 0;
            foreach (var filePath in fileList.filePaths)
            {
                if (!filePath.Contains(".docx"))
                {
                    Console.WriteLine("Not a .docx file, and therefore not converted");
                    continue;
                }
                var pdfPath = filePath.Replace(".docx", ".pdf");
                Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();

                Document doc = null;

                try
                {
                    doc = ap.Documents.Open(filePath, ReadOnly: false, Visible: false);
                    var wdFormatPdf = 17;
                    doc.SaveAs2(pdfPath, wdFormatPdf);
                    doc.Close();

                    textBoxDirPath.Text = textBoxDirPath.Text.Replace(filePath, filePath + " (check)");
                    toolStripProgressBar1.Increment(100 / fileList.filePaths.Count);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception Caught: " + ex.Message); // Could be that the document is already open (/) or Word is in Memory(?)
                    MessageBox.Show("Sorry, but there was an error in the application... \r\n " + ex.Data + "     " + ex.Message, "OOoooops... So Sorry!");
                }
                finally
                {
                    // Ambiguity between method 'Microsoft.Office.Interop.Word._Application.Quit(ref object, ref object, ref object)' and non-method 'Microsoft.Office.Interop.Word.ApplicationEvents4_Event.Quit'. Using method group.
                    // ap.Quit( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
                    ap.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ap);



                    if (doc != null)
                    {
                    }
                }
            }
            Cursor.Current = Cursors.Default;
            MessageBox.Show("Conversion Complete!");
        }


    }
}