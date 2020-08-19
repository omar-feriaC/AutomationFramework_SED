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
    class TestclsData
    {

        /// PRIVATE VARIABLE DECLARATION
        private static SLDocument _objFile;
        private static int _intCurrentRow;
        private static bool _blColumnNames;
        private static bool _blColNameSpaces;
        private static bool _blSheetInUse;
        private static int _intColCount;
        private static int _intRowCount;
        private static int _intRowStartIndex;
        private static int _intColumnStartIndex;
        private static int _intColumnRowIndex;
        private static Dictionary<string, int> _dicHeaders;


        /// METHODS
        public int CurrentRow
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

        public bool HasColumnNames
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

        public bool HasColumnSpaces
        {
            get { return _blColNameSpaces; }
            set { _blColNameSpaces = value; }
        }

        public int ColumnCount
        {
            get
            {
                return _intColCount;
            }
        }

        public int RowCount
        {
            get
            {
                return _intRowCount;
            }
        }

        public static int RowStartNo
        {
            get
            {
                return _intRowStartIndex;
            }
            set
            {
                _intRowStartIndex = value;
                _intCurrentRow = (_intCurrentRow > _intRowStartIndex ? _intCurrentRow : _intRowStartIndex);
            }
        }

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

        public static bool SheetInUse
        {
            get
            {
                return _blSheetInUse;
            }
        }


        /// CONSTRUCTOR
        public TestclsData()
        {
            _intCurrentRow = 2;
            _blColNameSpaces = false;
            _blColumnNames = true;
            _blSheetInUse = false;
            _intColCount = -1;
            _intRowCount = -1;
            _intRowStartIndex = 2;
            _intColumnStartIndex = 1;
            _intColumnRowIndex = 1;
        }

        /// PRIVATE METHODS

        private static void _GetRowCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intRowCount = _objFile.GetWorksheetStatistics().NumberOfRows > -1 ? _objFile.GetWorksheetStatistics().NumberOfRows : 0;
            }
        }

        private static void _GetColumnCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intColCount = _objFile.GetWorksheetStatistics().NumberOfColumns > -1 ? _objFile.GetWorksheetStatistics().NumberOfColumns : 0;
            }
        }
        
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

        private static string _RemoveSpecialCharsNoSpaces(string pstrString)
        {
            string strCleanString = pstrString;
            
strCleanString = strCleanString.Replace("\n", "");
            strCleanString = strCleanString.Replace("\r", "");

            return strCleanString;
        }

        public static void _GetHeadersNames()
        {
            _dicHeaders = new Dictionary<string, int>();
            if (!_blColNameSpaces)
            {
                for (int i = 1; i <= _objFile.GetWorksheetStatistics().NumberOfColumns; i++)
                {
                    _dicHeaders.Add(_RemoveSpecialChars(_objFile.GetCellValueAsString(1, i)), i);
                }
            }
            else
            {
                for (int i = 1; i <= _objFile.GetWorksheetStatistics().NumberOfColumns; i++)
                {
                    _dicHeaders.Add(_RemoveSpecialCharsNoSpaces(_objFile.GetCellValueAsString(1, i)), i);
                }
            }
        }

        public void LoadFile(string pstrFile, string pstrSheet)
        {
            if (!string.IsNullOrEmpty(pstrFile) && File.Exists(pstrFile))
            {
                //Load File
                _objFile = new SLDocument(pstrFile);
                _blColumnNames = true;
                _blSheetInUse = true;
                _intCurrentRow = 2;
                _intRowStartIndex = _objFile.GetWorksheetStatistics().StartRowIndex;
                _intColumnStartIndex = _objFile.GetWorksheetStatistics().StartColumnIndex;
                _intColumnRowIndex = 1;
                _GetRowCount();
                _GetColumnCount();
                _GetHeadersNames();

                //Load Sheet
                if (!string.IsNullOrEmpty(pstrSheet))
                {
                    if (_objFile.GetSheetNames().Contains(pstrSheet))
                    {
                        _objFile.SelectWorksheet(pstrSheet);
                    }
                }
            }
        }

        public string GetValueByIndex(int pintColumnIndex, [Optional] string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; };

            if (_objFile != null && _blSheetInUse)
            {
                pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, pintColumnIndex);
            }
            return pstrDefaultValue;
        }

        public string GetValueByColumn(string pstrColumnName, [Optional] string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; };

            if (_objFile != null && _blSheetInUse && _dicHeaders.Count > 0)
            {
                if (_dicHeaders.ContainsKey(pstrColumnName))
                {
                    pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, _dicHeaders[pstrColumnName]);
                }
                else
                {
                    pstrDefaultValue = "";
                }
            }
            return pstrDefaultValue;
        }

        public void SaveCellValue(string pstrColumnName, string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; };
            if (_dicHeaders.ContainsKey(pstrColumnName))
            {
                _objFile.SetCellValue(_intCurrentRow, _dicHeaders[pstrColumnName], pstrDefaultValue);
                _objFile.Save();
            }
        } 

    }
}
