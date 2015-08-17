using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using iasmcs;

namespace ExcellRunTest
{
    class WorksheetManager
    {
        public static int LEFT = 1;
        public static int CENTER = 2;
        public static int RIGHT = 4;
        public static int BOLD = 8;
        public static int NORMAL = 16;
        public static int ITALY = 32;
        public static int UNDERLINE = 64;
        public static int TOP = 128;
        public static int BOTTOM = 256;
        public static bool HORIZONTAL = true;
        public static bool VERTICAL = false;

        Excel.Worksheet Sheet;

        public WorksheetManager(Excel.Worksheet Sheet)
        {
            this.Sheet = Sheet;
        }

        #region

        public void SetWorksheetName(string WorksheetName)
        {
            Sheet.Name = WorksheetName;
        }
        #endregion


        #region 텍스트

        /// <summary>
        /// 해당 셀에 텍스트를 설정합니다.
        /// Ex) Text("A2", "Hello World!", WorkSheetManger.BOLD | WorkSheetManger.ITALY | WorkSheetManger.UNDERLINE, 11 );
        /// </summary>
        /// <param name="Cell"></param>
        /// <param name="Text"></param>
        /// <param name="Size"></param>
        public void InsertText(string Cell, string Text, int Setting = 5, int Size = 11, Color? TextColor = null)
        {
            Range TextRange = Sheet.get_Range(Cell);
            TextRange.Cells[1][1] = Text;
            TextRange.Font.Size = Size;
            TextRange.Font.Color = TextColor ?? Color.Black;
            if (Setting == BOLD)
            {
                TextRange.Font.Bold = true;
            }
            else if (Setting == ITALY)
            {
                TextRange.Font.Italic = true;
            }
            else if (Setting == UNDERLINE)
            {
                TextRange.Font.Underline = true;
            }
            else if (Setting == ( ITALY | BOLD ))
            {
                TextRange.Font.Bold = true;
                TextRange.Font.Italic = true;
            }
            else if (Setting == (ITALY | UNDERLINE))
            {
                TextRange.Font.Italic = true;
                TextRange.Font.Underline = true;
            }
            else if (Setting == (BOLD | UNDERLINE))
            {
                TextRange.Font.Italic = true;
                TextRange.Font.Underline = true;
            }
            else if (Setting == (BOLD | UNDERLINE | ITALY))
            {
                TextRange.Font.Underline = true;
                TextRange.Font.Italic = true;
                TextRange.Font.Bold = true;
            }
        }

        /// <summary>
        /// 해당 범위의 속성값을 변경합니다.
        /// Ex) ChangeTextSetting("A1", "C3", WorkSheetManger.BOLD | WorkSheetManager.ITALY);
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="Setting"></param>
        public void ChangeTextSetting(string start, string end, int Setting)
        {
            Range TextRange = Sheet.get_Range(start, end);

            if (Setting == BOLD)
            {
                TextRange.Font.Bold = true;
            }
            else if (Setting == ITALY)
            {
                TextRange.Font.Italic = true;
            }
            else if (Setting == UNDERLINE)
            {
                TextRange.Font.Underline = true;
            }
            else if (Setting == (ITALY | BOLD))
            {
                TextRange.Font.Bold = true;
                TextRange.Font.Italic = true;
            }
            else if (Setting == (ITALY | UNDERLINE))
            {
                TextRange.Font.Italic = true;
                TextRange.Font.Underline = true;
            }
            else if (Setting == (BOLD | UNDERLINE))
            {
                TextRange.Font.Italic = true;
                TextRange.Font.Underline = true;
            }
            else if (Setting == (BOLD | UNDERLINE | ITALY))
            {
                TextRange.Font.Underline = true;
                TextRange.Font.Italic = true;
                TextRange.Font.Bold = true;
            }
        }

        /// <summary>
        /// 해당 범위에 텍스트를 지정합니다. 범위를 벗어난 배열을 매개 변수로 지정하면 제대로 값이 지정되지 않을 수 있습니다.
        /// Ex) InsertRangeText("A1", "B4", ArrayText, WorksheetManager.HORIZONTAL);
        /// </summary>
        /// <param name="Start"></param>
        /// <param name="End"></param>
        /// <param name="Texts"></param>
        public void InsertRangeText(string Start, string End, string[,] Texts, bool Direction = true)
        {
            Range TextRange = Sheet.get_Range(Start, End);

            

            // 가로 순으로 그릴때
            if (Direction == HORIZONTAL)
            {

                for (int i = 0; i < Texts.GetLength(0); i++)
                {
                    for (int j = 0; j < Texts.GetLength(1); j++)
                    {
                        TextRange.Cells[j + 1][i + 1] = Texts[i, j];
                    }
                }
            }
            // 세로순으로 그릴때
            else if(Direction == VERTICAL)
            {

                for (int i = 0; i < Texts.GetLength(0); i++)
                {
                    for (int j = 0; j < Texts.GetLength(1); j++)
                    {
                        TextRange.Cells[i + 1][j + 1] = Texts[i, j];
                    }
                }
            }
        }
        #endregion


        #region 셀

        /// <summary>
        /// 해당 인덱스 범위 셀 병합
        /// Ex) Merge("A3","B4");
        /// </summary>
        /// <param name="Start"></param>
        /// <param name="End"></param>
        public void Merge(string Start, string End)
        {
            Range MergedRange = Sheet.get_Range(Start, End);
            MergedRange.Merge();
        }

        /// <summary>
        /// 지정된 범위의 정렬을 설정합니다. WorksheetManager.LEFT/WorksheetManager.CENTOR/Worksheet.RIGHT
        /// Ex) setAlign("A1", "B3", WorksheetManager.CENTOR);
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="WorkSheetAlign"></param>
        public void SetAlign(string start, string end, int WorkSheetAlign)
        {
            Excel.Range AlignRange = Sheet.get_Range(start, end);

            if (WorkSheetAlign == LEFT)
            {
                // 왼쪽 정렬
                AlignRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
            else if (WorkSheetAlign == CENTER)
            {
                // 가운데 정렬
                AlignRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                AlignRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else
            {
                // 오른쪽 정렬
                AlignRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            }
        }

        /// <summary>
        /// 컬럼의 너비를 설정합니다.
        /// Ex) setColumnsWidth("A",13.4);
        /// </summary>
        /// <param name="column"></param>
        /// <param name="width"></param>
        public void SetColumnsWidth(string column, double width)
        {
            Excel.Range ColRange = Sheet.get_Range(column + "1", column + "1");

            ColRange.EntireColumn.ColumnWidth = width;
        }

        /// <summary>
        /// 지정된 범위의 셀에 배경색을 지정합니다,
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="BackgroundColor"></param>
        public void SetBackgroundColor(string start, string end, Color BackgroundColor)
        {
            Range BackgroundRange = Sheet.get_Range(start, end);

            BackgroundRange.Interior.Color = BackgroundColor;
        }

        /// <summary>
        /// 범위 혹은 하나의 셀의 테두리를 지정합니다.
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="borders"></param>
        public void SetBorder(string start, string end, int borders = 511)
        {
            Range BorderRange = Sheet.get_Range(start, end);

            Borders border = BorderRange.Borders;

            if ((borders & TOP) == TOP)
            {
                border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            }
            if ((borders & LEFT) == LEFT)
            {
                border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            }
            if ((borders & RIGHT) == RIGHT)
            {
                border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            }
            if ((borders & BOTTOM) == BOTTOM)
            {
                border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }
        }

        /// <summary>
        /// 테이블 형식의 테두리를 설정합니다.
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public void SetTableBorder(string start, string end)
        {
            Range BorderRange = Sheet.get_Range(start, end);

            Borders border = BorderRange.Borders;

            border.LineStyle = XlLineStyle.xlContinuous;
        }

        /// <summary>
        /// 해당 범위의 컬럼의 너비를 지정합니다.
        /// Ex) SetRangColumnWidth("A","C",20);
        /// </summary>
        /// <param name="StartColumn"></param>
        /// <param name="EndColumn"></param>
        /// <param name="Width"></param>
        public void SetRangeColumnWidth(string StartColumn, string EndColumn, double Width)
        {
            Range ColumnRange = Sheet.get_Range(StartColumn + "1", EndColumn + "1");

            ColumnRange.EntireColumn.ColumnWidth = Width;
        }

        #endregion


        #region 이미지

        /// <summary>
        /// 해당 범위에 이미지를 삽입합니다.
        /// </summary>
        /// <param name="Start"></param>
        /// <param name="End"></param>
        /// <param name="FilePath"></param>
        public void InsertImage(string Start, string End, string FilePath)
        {
            Range ImageRange = Sheet.get_Range(Start, End);

            double ImageWidth = ImageRange.Columns.Width;
            double ImageHeight = ImageRange.Rows.Height;

            object missing = System.Reflection.Missing.Value;

            Pictures p = Sheet.Pictures(missing) as Pictures;
            Picture pic = null;

            pic = p.Insert(FilePath, missing);

            pic.ShapeRange.LockAspectRatio = MsoTriState.msoFalse;

            pic.Left = ImageRange.Left;
            pic.Top = ImageRange.Top;
            pic.ShapeRange.Width = ImageWidth.ToFloat();
            pic.ShapeRange.Height = ImageHeight.ToFloat();

        }

        #endregion

    }



}
