using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraRichEdit;
using DevExpress.XtraPrinting;

namespace ConvertWordToPDF
{
    public partial class MainForm : Form
    {
        enum COL_IDX { NAME, PATH }

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnAddFile_Click(object sender, EventArgs e)
        {
            string[] filePaths = FileDialogFunc.MultiFileFunc("파일 추가", "", "Word 파일 (*.doc, *.docx) | *.doc; *.docx;");
            if (filePaths.Length == 0)
                return;

            // 중복 경로 제거
            ExistFileList(ref filePaths);

            for (int i = 0; i < filePaths.Length; i++)
            {
                string path = filePaths[i];
                string name = Path.GetFileName(path);
                dgvMain.Rows.Add(new string[] { name, path });
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                for (int i = 0; i < dgvMain.Rows.Count; i++)
                {

                    string name = dgvMain.Rows[i].Cells[(int)COL_IDX.NAME].Value.ToString();
                    string path = dgvMain.Rows[i].Cells[(int)COL_IDX.PATH].Value.ToString();
                    if (path.Length == 0)
                        continue;

                    lbStatus.Text = string.Format($"작업중.. {name}");

                    string pdfPath = "";
                    if (!GetPDFFilePath(path, out pdfPath))
                    {
                        Debug.Assert(false);
                        continue;
                    }

                    using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
                    {
                        wordProcessor.LoadDocument(path);

                        PdfExportOptions options = new PdfExportOptions();
                        options.Compressed = false;
                        options.ImageQuality = PdfJpegImageQuality.Highest;

                        using (FileStream pdfFileStream = new FileStream(pdfPath, FileMode.Create))
                        {
                            wordProcessor.ExportToPdf(pdfFileStream, options);
                        }
                    }
                }

                lbStatus.Text = "완료";
                MessageBox.Show("변환 완료", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                lbStatus.Text = "오류";
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void ExistFileList(ref string[] filePaths)
        {
            try
            {
                List<string> tmpPathList = filePaths.ToList();
                List<int> existInd = new List<int>();
                for (int i = 0; i < dgvMain.Rows.Count; i++)
                {
                    string path = dgvMain.Rows[i].Cells[(int)COL_IDX.PATH].Value.ToString();
                    if (path.Length == 0)
                        continue;

                    if (tmpPathList.Any(c => string.Compare(path, c, true) == 0))
                        existInd.Add(tmpPathList.FindLastIndex(c => string.Compare(path, c, true) == 0));
                }

                existInd.Sort();

                for (int i = existInd.Count - 1; i >= 0; i--)
                    tmpPathList.RemoveAt(existInd[i]);

                filePaths = tmpPathList.ToArray();
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                Debug.Assert(false);
            }
        }

        private bool GetPDFFilePath(string wordpath, out string pdfpath)
        {
            pdfpath = "";

            try
            {
                string withoutExtensionName = Path.GetFileNameWithoutExtension(wordpath);
                string dir = Path.GetDirectoryName(wordpath);
                string name = string.Format($"{withoutExtensionName}.PDF");
                pdfpath = Path.Combine(dir, name);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                Debug.Assert(false);
                return false;
            }

            return true;
        }
    }
}
