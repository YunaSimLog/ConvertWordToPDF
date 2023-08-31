using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertWordToPDF
{
    public class FileDialogFunc
    {
        public static string SingleFileFunc(string strTitle, string strFileName, string strFilter)
        {
            string strFilePath = "";
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = strTitle;
                ofd.FileName = strFileName;
                ofd.Filter = strFilter;
                ofd.Multiselect = false;

                //파일 오픈창 로드
                DialogResult dr = ofd.ShowDialog();

                //OK버튼 클릭시
                if (dr == DialogResult.OK)
                    strFilePath = ofd.FileName;
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                Debug.Assert(false);
            }

            return strFilePath;
        }

        public static string[] MultiFileFunc(string strTitle, string strFileName, string strFilter)
        {
            string[] strarrFilePath = null;
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = strTitle;
                ofd.FileName = strFileName;
                ofd.Filter = strFilter;
                ofd.Multiselect = true;

                //파일 오픈창 로드
                DialogResult dr = ofd.ShowDialog();

                //OK버튼 클릭시
                if (dr == DialogResult.OK)
                    strarrFilePath = ofd.FileNames;
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                Debug.Assert(false);
            }

            return strarrFilePath;
        }
    }
}
