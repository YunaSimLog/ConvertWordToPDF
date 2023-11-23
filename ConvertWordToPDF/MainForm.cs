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
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;

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
            if (filePaths == null || filePaths.Length == 0)
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

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                for (int i = 0; i < dgvMain.Rows.Count; i++)
                {

                    string name = dgvMain.Rows[i].Cells[(int)COL_IDX.NAME].Value.ToString();
                    string path = dgvMain.Rows[i].Cells[(int)COL_IDX.PATH].Value.ToString();
                    if (path.Length == 0)
                        continue;

                    lbStatus.Text = string.Format($"작업중.. {name}");
                    dgvMain.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(255, 227, 121);

                    string pdfPath = "";
                    if (!GetPDFFilePath(path, out pdfPath))
                    {
                        Debug.Assert(false);
                        continue;
                    }

                    try
                    {
                        object filename = path;
                        Document doc = word.Documents.Open(ref filename, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        doc.Activate();

                        object outputFileName = pdfPath;
                        object fileFormat = WdSaveFormat.wdFormatPDF;

                        doc.SaveAs(ref outputFileName,
                            ref fileFormat, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                        ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                        doc = null;
                    }
                    catch (Exception ex)
                    {
                        dgvMain.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(229, 37, 53);
                        lbStatus.Text = "오류";
                        Debug.Print(ex.Message);
                        Debug.Assert(false);
                    }

                    dgvMain.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(69, 196, 176);
                }

                lbStatus.Text = "완료";
                MessageBox.Show("변환 완료", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                word = null;
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

        private void btnDelete_Click(object sender, EventArgs e)
        {
            List<int> selLows = new List<int>();
            foreach (DataGridViewRow row in dgvMain.SelectedRows)
                selLows.Add(row.Index);

            selLows.Sort();

            for (int i = selLows.Count - 1; i >= 0; i--)
                dgvMain.Rows.RemoveAt(selLows[i]);
        }

        private void btnDeleteAll_Click(object sender, EventArgs e)
        {
            dgvMain.Rows.Clear();
            dgvMain.Refresh();
        }

        private void dgvMain_DragDrop(object sender, DragEventArgs e)
        {
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (filePaths.Length < 1)
                return;

            // Doc 문서 추리기
            OnlyDocFiles(ref filePaths);

            // 중복 경로 제거
            ExistFileList(ref filePaths);

            for (int i = 0; i < filePaths.Length; i++)
            {
                string path = filePaths[i];
                string name = Path.GetFileName(path);
                dgvMain.Rows.Add(new string[] { name, path });
            }
        }

        private void dgvMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
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

        private void OnlyDocFiles(ref string[] filePaths)
        {
            List<string> tmpPathList = new List<string>();

            for (int i = 0; i < filePaths.Length; i++)
            {
                string path = filePaths[i];
                string extension = Path.GetExtension(path);

                // word문서 아닐 경우 예외처리
                if (string.Compare(extension, ".doc", true) != 0 && string.Compare(extension, ".docx", true) != 0)
                    continue;

                tmpPathList.Add(path);
            }

            filePaths = tmpPathList.ToArray();
        }
    }
}
