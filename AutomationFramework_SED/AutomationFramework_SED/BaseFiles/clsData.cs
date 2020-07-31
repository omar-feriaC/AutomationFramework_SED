using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.BaseFiles
{
    class clsData
    {
        /// PRIVATE VARIABLE DECLARATION
        private static SLDocument _objFile;
        private static int  _intCurrentRow;
        private static bool _blColumnNames;
        private static bool _blSheetInUse;
        private static int  _intColCount;
        private static int  _intRowCount;
        private static int  _intRowStartIndex;
        private static int  _intColumnStartIndex;
        private static int  _intColumnRowIndex;
        private static Dictionary<string,string> _dicHeaders;



        private static string _strDataFile;


        public static string DataFile
        {
            get { return _strDataFile; }
            set { DataFile = _strDataFile; }
        }
    



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

        public static int ColumnCount
        {
            get
            {
                return _intColCount;
            }
        }

        public static int RowCount
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
                //_intColumnRowIndex = (_blColumnNames ? value - 1 : value);
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



        static clsData()
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

        private static void _GetColumnCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intColCount = _objFile.GetWorksheetStatistics().NumberOfColumns;
                _intColCount = (_intColCount >= 0 ? _intColCount : 0);
            }
        }

        private static void _GetRowCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intRowCount = _objFile.GetWorksheetStatistics().NumberOfRows;
                _intRowCount = (_intRowCount >= 0 ? _intRowCount : 0);
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

        public static void _GetHeadersNames()
        {
            _dicHeaders = new Dictionary<string,string> ();
            for (int i = 1; i<= _objFile.GetWorksheetStatistics().NumberOfColumns; i++)
            {
                _dicHeaders.Add(_objFile.GetCellValueAsString(1,i), _objFile.GetCellValueAsString(1, i));
            }
        }

        public static bool _ColumnExists(string pstrColumnName)
        {
            if (_dicHeaders != null)
            {
                return _dicHeaders.ContainsKey(pstrColumnName) ? true : false;
            }
            else
            {
                return false;
            }
        }



    }
}
