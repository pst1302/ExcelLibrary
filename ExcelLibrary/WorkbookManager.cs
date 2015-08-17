using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using iasmcs;

namespace ExcellRunTest
{
    class WorkbookManager
    {
        Workbook WorkBook = null;
        int NumberOfWorksheet;
        Worksheet[] worksheets = null;

        bool Is2003 = false;

        /// <summary>
        /// Excel Application을 받아 지정된 갯수의 Worksheet를 가진 워크북을 생성합니다.
        /// </summary>
        /// <param name="ExcelApp"></param>
        /// <param name="NumOfWorksheet"></param>
        public WorkbookManager(Application ExcelApp, int NumOfWorksheet)
        {
            CheckVersion(ExcelApp);

            WorkBook = ExcelApp.Workbooks.Add();

            MakeWorksheet(NumOfWorksheet);

            NumberOfWorksheet = NumOfWorksheet;
        }

        // Worksheet 생성
        private void MakeWorksheet(int NumOfWorksheet)
        {
            worksheets = new Worksheet[NumOfWorksheet];

            for (int i = 0; i < worksheets.Length; i++)
            {
                worksheets[i] = WorkBook.Worksheets.Add();
            }
        }

        // 버전 체크 : 2003과 2003 이외의 버전 구분
        private void CheckVersion(Application ExcelApp)
        {
            string curExcelVer = ExcelApp.Version.Substring(0, 2);

            if (curExcelVer == "11")
            {
                Is2003 = true;
            }
            else
            {
                Is2003 = false;
            }
        }

        /// <summary>
        /// 해당 인덱스의 Worksheet를 가져옵니다.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Worksheet GetWorksheets(int index)
        {
            return worksheets[NumberOfWorksheet - index - 1];
        }

        /// <summary>
        /// 현재 Workbook을 파일로 저장합니다.
        /// </summary>
        /// <param name="FileName"></param>
        public void SaveFile(string FilePath)
        {
            string excutor;

            if (Is2003)
            {
                excutor = ".xls";
            }
            else
            {
                excutor = ".xlsx";
            }

            // 만약 파일이 있으면 삭제
            if (iasm.File.FileExist(FilePath + excutor))
                iasm.File.DeleteFile(FilePath + excutor);


            WorkBook.SaveAs(FilePath + excutor);
            WorkBook.Close(true);
        }

    }
}
