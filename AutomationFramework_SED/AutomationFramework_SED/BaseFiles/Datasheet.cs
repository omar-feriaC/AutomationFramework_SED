using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace AutomationFramework_SED.BaseFiles
{
    class Datasheet
    {

        /// PRIVATE VARIABLE DECLARATION
        private static SLDocument _objFile;
        private static int _intCurrentRow;
        private static bool _blColumnNames;
        private static bool _blSheetInUse;
        private static int _intColCount;
        private static int _intRowCount;
        private static int _intRowStartIndex;
        private static int _intColumnStartIndex;
        private static int _intColumnRowIndex;

        /// <summary>
        /// Indicates the current row we are using in the file. It can't be less that RowStartNo. By default is 2.
        /// </summary>
        public static int CurrentRow
        {
            get
            {
                return _intCurrentRow;
            }
            set
            {
                _intCurrentRow = (value > _intRowStartIndex ? value : _intRowStartIndex);
            }
        }
        /// <summary>
        /// Indicates if the file has column names on the given row (by default the first).
        /// </summary>
        public static bool HasColumnNames
        {
            get
            {
                return _blColumnNames;
            }
            set
            {
                _blColumnNames = value;
                _intRowStartIndex = (value ? 2 : 1);
            }
        }
        /// <summary>
        /// Indicates the number of columns in the file. By default is TRUE.
        /// </summary>
        public static int ColumnCount
        {
            get
            {
                return _intColCount;
            }
        }
        /// <summary>
        /// Indicates the number of rows in the file.
        /// </summary>
        public static int RowCount
        {
            get
            {
                return _intRowCount;
            }
        }
        /// <summary>
        /// Indicates the index of the starting row in the file. By default is 2.
        /// </summary>
        public static int RowStartNo
        {
            get
            {
                return _intRowStartIndex;
            }
            set
            {
                _intRowStartIndex = value;
                //_intColumnRowIndex = (_blColumnNames ? value - 1 : value);
                _intCurrentRow = (_intCurrentRow > _intRowStartIndex ? _intCurrentRow : _intRowStartIndex);
            }
        }
        /// <summary>
        /// Indicates the index of the starting column in the file. By default is 1.
        /// </summary>
        public static int ColumnStartNo
        {
            get
            {
                return _intColumnStartIndex;
            }
            set
            {
                _intColumnStartIndex = value;
            }
        }
        /// <summary>
        /// Indicates the index of the row containing the headers. By default is 1.
        /// </summary>
        public static int ColumnNameRow
        {
            get
            {
                return _intColumnRowIndex;
            }
            set
            {
                _intColumnRowIndex = value;
                _intRowStartIndex = (_blColumnNames ? value + 1 : value);
                _intCurrentRow = (_intCurrentRow > _intRowStartIndex ? _intCurrentRow : _intRowStartIndex);
                _GetColumnCount();
            }
        }
        /// <summary>
        /// Indicates if we have a sheet in use.
        /// </summary>
        public static bool SheetInUse
        {
            get
            {
                return _blSheetInUse;
            }
        }

        /// CONSTRUCTOR

        /// <summary>
        /// Sets the default values for the properties.
        /// </summary>
        static Datasheet()
        {
            _intCurrentRow = 2;
            _blColumnNames = true;
            _blSheetInUse = false;
            _intColCount = -1;
            _intRowCount = -1;
            _intRowStartIndex = 2;
            _intColumnStartIndex = 1;
            _intColumnRowIndex = 1;
        }

        /// PRIVATE METHODS

        private static void _GetColumnCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intColCount = 0;
                for (int i = _intColumnStartIndex; i <= _objFile.GetWorksheetStatistics().EndColumnIndex; i++)
                {
                    if (_objFile.HasCellValue(_intColumnRowIndex, i))
                    {
                        _intColCount++;
                    }
                }
            }
        }

        private static void _GetRowCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intRowCount = 0;
                for (int i = _intRowStartIndex; i <= _objFile.GetWorksheetStatistics().EndRowIndex; i++)
                {
                    if (_objFile.HasCellValue(i, 1))
                    {
                        _intRowCount++;
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strColumnName"></param>
        /// <returns></returns>
        private static int _GetColumnIndex(string pstrColumnName)
        {
            int intCol = -1;
            if (_objFile != null && _blSheetInUse)
            {
                for (intCol = _intColumnStartIndex; intCol <= _intColCount; intCol++)
                {
                    if (_RemoveSpecialChars(_objFile.GetCellValueAsString(_intColumnRowIndex, intCol)).Equals(_RemoveSpecialChars(pstrColumnName)))
                    {
                        break;
                    }
                }
            }
            return intCol;
        }

        private static string _RemoveSpecialChars(string pstrString)
        {
            string strCleanString = pstrString;

            strCleanString = strCleanString.Replace(" ", "");
            strCleanString = strCleanString.Replace("\n", "");
            strCleanString = strCleanString.Replace("\r", "");

            return strCleanString;
        }

        public static void FileOpen(string pstrFile, [Optional] string pstrSheet)
        {
            if (!string.IsNullOrEmpty(pstrFile))
            {
                if (File.Exists(pstrFile))
                {
                    _objFile = new SLDocument(pstrFile);
                    _blColumnNames = true;
                    _intCurrentRow = 2;
                    _intRowStartIndex = _objFile.GetWorksheetStatistics().StartRowIndex;
                    _intColumnStartIndex = _objFile.GetWorksheetStatistics().StartColumnIndex;
                    _intColumnRowIndex = 1;
                    _blSheetInUse = true;
                    _GetRowCount();
                    _GetColumnCount();

                    if (!string.IsNullOrEmpty(pstrSheet))
                    {
                        if (_objFile.GetSheetNames().Contains(pstrSheet))
                        {
                            _objFile.SelectWorksheet(pstrSheet);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Get the value from the Sheet in use.
        /// Ensure to set the CurrentRow property.
        /// Ensure to have a SheetInUse through the property SheetInUse.
        /// Otherwise select a sheet.
        /// </summary>
        /// <param name="pstrColumnName"></param>
        /// <returns></returns>
        public static string GetValue(string pstrColumnName, [Optional] string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; };
            int intColumnIndex = 0;

            if (_objFile != null && _blSheetInUse)
            {
                intColumnIndex = _GetColumnIndex(pstrColumnName);
                pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, intColumnIndex);
            }
            return pstrDefaultValue;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pintColumnIndex"></param>
        /// <param name="pstrDefaultValue"></param>
        /// <returns></returns>
        public static string GetValue(int pintColumnIndex, [Optional] string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; };

            if (_objFile != null && _blSheetInUse)
            {
                pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, pintColumnIndex);
            }
            return pstrDefaultValue;
        }



        }
}
