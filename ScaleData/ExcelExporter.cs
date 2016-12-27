using Microsoft.Office.Interop.Excel;
using ScalesData.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ScalesData
{
    class ExcelExporter
    {
        private const int COL_INDEX = 1;
        private const int COL_TIME_INDEX = 2;
        private const int INVALID_VALUE = -1;

        private const int TEXT_NORMAL = 14;
        private const int TEXT_SMALL = 12;

        private Dictionary<Object, int> mIndexMapper = new Dictionary<object, int>();
        private int mStartRow = 7;
        private int mEndRow = 7;
        private int mStartCol = 1;
        private int mEndCol = 1;

        private Application mExcel;
        private _Workbook mWorkbook;
        private _Worksheet mSheet;

        public DateTime StartDate { get; set; }

        public DateTime EndDate { get; set; }

        public List<ScalesValue> DataList { get; set; }
        public string[] ScalesList { get; set; }
        public string Password { get; set; }

        public void export()
        {
            mExcel = new Application();
            mExcel.ScreenUpdating = false;
            mWorkbook = mExcel.Workbooks.Add("");
            mSheet = mWorkbook.ActiveSheet;

            mSheet.Cells.Font.Size = TEXT_NORMAL;
            mSheet.Cells.Font.Name = "Times New Roman";

            addTableHeader();
            addTableContent();
            addForm();

            protect();
            saveAndClose();
        }

        private void addTableHeader()
        {
            Range countRange = getRange(mStartRow, ++mEndRow, COL_INDEX, COL_INDEX);
            countRange.Merge();
            countRange.Value = Resources.TITLE_COUNT;
            setCellData(mEndRow, COL_TIME_INDEX, Resources.TITLE_TIME);

            if (ScalesList != null)
            {
                for (int i = 1; i <= ScalesList.Length; i++)
                {
                    int col = COL_TIME_INDEX + i;
                    string tag = ScalesList[i - 1];
                    mIndexMapper.Add(tag, col);
                    setCellData(mEndRow, col, tag);
                }
            }

            Range header = getRange(mStartRow, mEndRow, mStartCol, mEndCol);
            header.Font.Bold = true;
            header.Font.Size = TEXT_SMALL;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            header.VerticalAlignment = XlVAlign.xlVAlignCenter;

            Range material = getRange(mStartRow, mStartRow, COL_INDEX + 1, mEndCol);
            material.Merge();
            material.Value = Resources.TITLE_MATERIAL;
            material.Font.Size = TEXT_NORMAL;
        }

        private void addTableContent()
        {
            for (int i = 1; i <= DataList.Count; i++)
            {
                ScalesValue data = DataList[i - 1];
                string date = data.Time.ToString();
                int row_index, col_index;

                if (!mIndexMapper.TryGetValue(date, out row_index))
                {
                    row_index = mEndRow + 1;
                    mIndexMapper.Add(date, row_index);
                }
                if (!mIndexMapper.TryGetValue(data.ScaleTag, out col_index))
                {
                    col_index = mEndCol + 1;
                    mIndexMapper.Add(data.ScaleTag, col_index);
                }

                setCellData(row_index, COL_TIME_INDEX, date);
                setCellData(row_index, col_index, data.Weight);
                setCellData(row_index, COL_INDEX, i);

                for (int j = COL_TIME_INDEX + 1; j <= mEndCol; j++)
                {
                    Range cell = mSheet.Cells[row_index, j];
                    if (cell.Value == null)
                    {
                        cell.Value = 0;
                    }
                }
            }

            // Hardcode time column
            Range timeRange = getRange(mStartRow, mEndRow, COL_TIME_INDEX, COL_TIME_INDEX);
            timeRange.EntireColumn.AutoFit();
            Range countRange = getRange(mStartRow, mEndRow, COL_INDEX, COL_INDEX);
            countRange.Columns.AutoFit();
            Range tableRange = getRange(mStartRow, mEndRow, mStartCol, mEndCol);
            Borders border = tableRange.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
        }

        private void addForm()
        {
            DateTime date = StartDate;
            if (DataList != null && DataList.Count > 0)
            {
                date = DataList[0].Time;
            }
            if (date.Year == 1)     // This date is still auto created. So set date to current date
            {
                date = DateTime.Now;
            }

            object[,] form =
            {
                { Resources.TITLE , null, null, null, null, null, null, null, null, null, null, null},
                { Resources.SHIFT, null, Resources.DATE, date.Day, Resources.MONTH, date.Month, Resources.YEAR, date.Year, null, null, null, null},
                { Resources.TYPE, null, null, null, null, null, null, null, null, null, null, null},
                { Resources.LEADER_NAME, null, null, Resources.KCS, null, null, null, null, null, null, null, null},
                { Resources.BATCH, null, null, Resources.PRODUCTION_SMOOTHING, null, null, null, null, null, null, null, null}
            };

            int row = form.GetLength(0) + 1;
            int col = Math.Max(form.GetLength(1), mEndCol);
            Range formRange = getRange(1, form.GetLength(0), 1, form.GetLength(1));
            formRange.Value = form;

            Range unit = getRange(row, row, col - 2, col);
            unit.Merge();
            unit.Font.Italic = true;
            unit.Value = Resources.UNIT;

            Range title = getRange(1, 1, 1, col);
            title.Merge();
            title.Font.Bold = true;
            title.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private void protect()
        {
            Range tableRange = mSheet.Range[mSheet.Cells[mStartRow, mStartCol], mSheet.Cells[mEndRow, mEndCol]];
            mSheet.Cells.Locked = false;
            tableRange.Locked = true;
            mSheet.Protect(Password, AllowFormattingCells: true, AllowFormattingColumns:true, AllowFormattingRows: true);
        }

        private void saveAndClose()
        {
            string filename = getFileName();
            try
            {
                mWorkbook.SaveAs(filename);
            }
            catch (COMException e)
            {
                if ((uint)e.ErrorCode == 0x800A03EC)  // File exists err
                {
                    filename = addFileNameUnique(filename);
                    mWorkbook.SaveAs(filename);
                }
            }

            mWorkbook.Close();
            MessageBoxEx.Show("Report created at " + filename + ".");
        }

        private Range getRange(int startRow, int endRow, int startCol, int endCol)
        {
            return mSheet.Range[mSheet.Cells[startRow, startCol], mSheet.Cells[endRow, endCol]];
        }

        private void setCellData(int row, int col, object data)
        {
            mSheet.Cells[row, col] = data;
            if (row > mEndRow)
            {
                mEndRow = row;
            }
            if (col > mEndCol)
            {
                mEndCol = col;
            }
        }

        private string getFileName()
        {
            string filename = "{0}\\Report_{1}_{2}.xlsx";
            filename = String.Format(filename,
                Settings.Default.Destination,
                StartDate != null ? StartDate.ToString("MM.dd.yyyy") : "",
                EndDate != null ? EndDate.ToString("MM.dd.yyyy") : "");

            return filename;
        }

        private string addFileNameUnique(string _filename)
        {
            string directoryPath = Path.GetDirectoryName(_filename);
            string filename = Path.GetFileNameWithoutExtension(_filename);
            string extension = Path.GetExtension(_filename);
            string format = directoryPath + "\\" + filename + " ({0})." + extension;
            string path;
            int i = 1;

            while (true)
            {
                path = string.Format(format, i++);
                if (!File.Exists(path))
                {
                    break;
                }
            }

            return path;
        }
    }
}
