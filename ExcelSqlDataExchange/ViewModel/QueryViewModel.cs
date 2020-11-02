using System;
using System.Collections.Generic;
using ExcelSqlDataExchange.Support;
using Dapper;
using ExcelSqlDataExchange.Models;
using System.Data;
using System.Windows;
using ClosedXML.Excel;
using System.Linq;
using System.Collections.ObjectModel;

namespace ExcelSqlDataExchange.ViewModel
{
    class QueryViewModel : ModelView
    {
        public ModelView modelView { get; set; }
        List<Equipment> Equipmentlist = new List<Equipment>();
        string columnFilter = string.Empty;

        #region Export Option
        private string _exportOption = "Export Data by User Input";
        public string ExportOption
        {
            get { return _exportOption; }
            set
            {
                _exportOption = value;
                SetPropertyChanged("ExportOption");
            }
        }

        private string _searchKeyWord= "Building";
        public string SearchKeyWord
        {
            get { return _searchKeyWord; }
            set
            {
                _searchKeyWord = value;
                SetPropertyChanged("SearchKeyWord");
            }
        }

        private string _searchValue;
        public string SerachValue
        {
            get { return _searchValue; }
            set
            {
                _searchValue = value;
                SetPropertyChanged("SerachValue");
            }
        }

        private bool _selectAll;
        public bool SelectAll
        {
            get { return _selectAll; }
            set
            {
                if (value)
                {
                    SelectNone = false;

                    IDCheck = true;
                    NameCheck=true;
                    TypeCheck = true;
                    BarCodeCheck = true;

                    BuildingCheck = true;
                    LevelCheck = true;
                    RoomCheck = true;
                    ZoneCheck = true;

                    DocLinkCheck = true;
                    PhotoLinkCheck = true;

                    ClassificationCheck = true;
                    MaterialTypeCheck = true;
                    ConsequencePriorityCheck = true;
                    OperationStatusCheck = true;

                    ManufacturerCheck = true;
                    YearCheck = true;
                    DegradationInforCheck = true;
                    DetailCheck = true;

                    InspectionStatusCheck = true;
                    AlarmTypeCheck = true;
                    CollectedByCheck = true;
                    CollectedOnCheck = true;
                    NotesCheck = true;
                    InspectPhotoCheck = true;
                    AttachmentLinkCheck = true;
                }

                _selectAll = value;
                SetPropertyChanged("SelectAll");
            }
        }

        private bool _selectNone;
        public bool SelectNone
        {
            get { return _selectNone; }
            set
            {
                if (value)
                {
                    SelectAll = false;

                    IDCheck = false;
                    NameCheck = false;
                    TypeCheck = false;
                    BarCodeCheck = false;

                    BuildingCheck = false;
                    LevelCheck = false;
                    RoomCheck = false;
                    ZoneCheck = false;

                    DocLinkCheck = false;
                    PhotoLinkCheck = false;

                    ClassificationCheck = false;
                    MaterialTypeCheck = false;
                    ConsequencePriorityCheck = false;
                    OperationStatusCheck = false;

                    ManufacturerCheck = false;
                    YearCheck = false;
                    DegradationInforCheck = false;
                    DetailCheck = false;

                    InspectionStatusCheck = false;
                    AlarmTypeCheck = false;
                    CollectedByCheck = false;
                    CollectedOnCheck = false;
                    NotesCheck = false;
                    InspectPhotoCheck = false;
                    AttachmentLinkCheck = false;
                }
                    
                _selectNone = value;
                SetPropertyChanged("SelectNone");
            }
        }
        #endregion

        #region Excel
        private ObservableCollection<Equipment> _dataGrid = new ObservableCollection<Equipment>();
        public ObservableCollection<Equipment> DataGrid
        {
            get { return _dataGrid; }
            set
            {
                _dataGrid = value;
                SetPropertyChanged("DataGrid");
            }
        }
        #endregion

        #region User Input

        #region General Infor

        private bool _id ;
        public bool IDCheck
        {
            get { return _id; }
            set
            {
                _id = value;
                SetPropertyChanged("IDCheck");
            }
        }

        private bool _name;
        public bool NameCheck
        {
            get { return _name; }
            set
            {
                _name = value;
                SetPropertyChanged("NameCheck");
            }
        }

        private bool _type;
        public bool TypeCheck
        {
            get { return _type; }
            set
            {
                _type = value;
                SetPropertyChanged("TypeCheck");
            }
        }

        private bool _barCode;
        public bool BarCodeCheck
        {
            get { return _barCode; }
            set
            {
                _barCode = value;
                SetPropertyChanged("BarCodeCheck");
            }
        }
        #endregion

        #region Location
        private bool _building;
        public bool BuildingCheck
        {
            get { return _building; }
            set
            {
                _building = value;
                SetPropertyChanged("BuildingCheck");
            }
        }

        private bool _level;
        public bool LevelCheck
        {
            get { return _level; }
            set
            {
                _level = value;
                SetPropertyChanged("LevelCheck");
            }
        }

        private bool _room;
        public bool RoomCheck
        {
            get { return _room; }
            set
            {
                _room = value;
                SetPropertyChanged("RoomCheck");
            }
        }

        private bool _zone;
        public bool ZoneCheck
        {
            get { return _zone; }
            set
            {
                _zone = value;
                SetPropertyChanged("ZoneCheck");
            }
        }


        #endregion

        #region Documentation
        private bool _docLink;
        public bool DocLinkCheck
        {
            get { return _docLink; }
            set
            {
                _docLink = value;
                SetPropertyChanged("DocLinkCheck");
            }
        }

        private bool _docPhotoLink;
        public bool PhotoLinkCheck
        {
            get { return _docPhotoLink; }
            set
            {
                _docPhotoLink = value;
                SetPropertyChanged("PhotoLinkCheck");
            }
        }

        #endregion

        #region Classification
        private bool _classification;
        public bool ClassificationCheck
        {
            get { return _classification; }
            set
            {
                _classification = value;
                SetPropertyChanged("ClassificationCheck");
            }
        }

        private bool _materialType;
        public bool MaterialTypeCheck
        {
            get { return _materialType; }
            set
            {
                _materialType = value;
                SetPropertyChanged("MaterialTypeCheck");
            }
        }

        private bool _consequencePriority;
        public bool ConsequencePriorityCheck
        {
            get { return _consequencePriority; }
            set
            {
                _consequencePriority = value;
                SetPropertyChanged("ConsequencePriorityCheck");
            }
        }

        private bool _opeationStatus;
        public bool OperationStatusCheck
        {
            get { return _opeationStatus; }
            set
            {
                _opeationStatus = value;
                SetPropertyChanged("OperationStatusCheck");
            }
        }

        #endregion

        #region Manufacturer
        private bool _manufacturer;
        public bool ManufacturerCheck
        {
            get { return _manufacturer; }
            set
            {
                _manufacturer = value;
                SetPropertyChanged("ManufacturerCheck");
            }
        }

        private bool _mYear;
        public bool YearCheck
        {
            get { return _mYear; }
            set
            {
                _mYear = value;
                SetPropertyChanged("YearCheck");
            }
        }

        private bool _degradationInfo;
        public bool DegradationInforCheck
        {
            get { return _degradationInfo; }
            set
            {
                _degradationInfo = value;
                SetPropertyChanged("DegradationInforCheck");
            }
        }

        private bool _detail;
        public bool DetailCheck
        {
            get { return _detail; }
            set
            {
                _detail = value;
                SetPropertyChanged("DetailCheck");
            }
        }
        #endregion

        #region Inspection
        private bool _status;
        public bool InspectionStatusCheck
        {
            get { return _status; }
            set
            {
                _status = value;
                SetPropertyChanged("InspectionStatusCheck");
            }
        }

        private bool _alarmType;
        public bool AlarmTypeCheck
        {
            get { return _alarmType; }
            set
            {
                _alarmType = value;
                SetPropertyChanged("AlarmTypeCheck");
            }
        }

        private bool _collectedBy;
        public bool CollectedByCheck
        {
            get { return _collectedBy; }
            set
            {
                _collectedBy = value;
                SetPropertyChanged("CollectedByCheck");
            }
        }

        private bool _collectedOn;
        public bool CollectedOnCheck
        {
            get { return _collectedOn; }
            set
            {
                _collectedOn = value;
                SetPropertyChanged("CollectedOnCheck");
            }
        }

        private bool _notes;
        public bool NotesCheck
        {
            get { return _notes; }
            set
            {
                _notes = value;
                SetPropertyChanged("NotesCheck");
            }
        }

        private bool _inspectionPhotoLink ;
        public bool InspectPhotoCheck
        {
            get { return _inspectionPhotoLink; }
            set
            {
                _inspectionPhotoLink = value;
                SetPropertyChanged("InspectPhotoCheck");
            }
        }

        private bool _attachmentLink ;
        public bool AttachmentLinkCheck
        {
            get { return _attachmentLink; }
            set
            {
                _attachmentLink = value;
                SetPropertyChanged("AttachmentLinkCheck");
            }
        }
        #endregion
        #endregion

        #region DelegateCommand

        private DelegateCommand _export;
        public DelegateCommand Export
        {
            get { return _export; }
            set
            {
                _export = value;
                SetPropertyChanged("Export");
            }
        }

        #endregion

        #region Window
        private Guid _viewId;

        public Guid ViewID
        {
            get { return _viewId; }
        }

        #endregion

        public QueryViewModel()
        {
            modelView = new ModelView();
            Export = new DelegateCommand(ExportAction);
            DataGrid.Clear();
            _viewId = Guid.NewGuid();
        }

        private void ExportAction()
        {
            try
            {
                Equipmentlist.Clear();
            using (IDbConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVal("AssetDB")))
            {
                if (ExportOption == "Export all the Data")
                {
                    Equipmentlist = connection.Query<Equipment>("dbo.Get_EquipmentList").ToList();
                    SaveToExcel();
                    LinkToDataGrid();
                    //Equipmentlist = connection.Query < Equipment >("dbo.Get_EquipmentList @equipmentId", new { equipmentId = "1" }).ToList();}
                }
                else {
                    FilterBySelection(connection);
                    SaveToExcelByFilter();
                }
            }
            MessageBox.Show("The data has been Export to a Excel File!");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                WindowManager.CloseWindow(ViewID);
                throw;
            }
        }


        private void LinkToDataGrid()
        {
            DataGrid.Clear();
            foreach (var equipment in Equipmentlist)
                DataGrid.Add(equipment);
        }
        private void SaveToExcel()
        {
            var xlWorkbook = new XLWorkbook();
            var xlSheet = xlWorkbook.Worksheets.Add("Equipment List");

        #region Set Up
            xlSheet.Cell(1, 1).Value = "Genreal Information";
            xlSheet.Cell(2, 1).Value = "ID";
            xlSheet.Cell(2, 2).Value = "Name";
            xlSheet.Cell(2, 3).Value = "System";
            xlSheet.Cell(2, 4).Value = "Type";

            xlSheet.Cell(1, 5).Value = "Location";
            xlSheet.Cell(2, 5).Value = "Building";
            xlSheet.Cell(2, 6).Value = "Level";
            xlSheet.Cell(2, 7).Value = "Room";
            xlSheet.Cell(2, 8).Value = "Zone";

            xlSheet.Cell(1, 9).Value = "Documentation";
            xlSheet.Cell(2, 9).Value = "Doc Link";
            xlSheet.Cell(2, 10).Value = "Photo Link";

            xlSheet.Cell(1, 11).Value = "Classification";
            xlSheet.Cell(2, 11).Value = "Classification";
            xlSheet.Cell(2, 12).Value = "Material Type";
            xlSheet.Cell(2, 13).Value = "Consequence Priority";
            xlSheet.Cell(2, 14).Value = "OpeationStatus";

            xlSheet.Cell(1, 15).Value = "Manufacturer";
            xlSheet.Cell(2, 15).Value = "Manufacturer";
            xlSheet.Cell(2, 16).Value = "Year";
            xlSheet.Cell(2, 17).Value = "Degradation Info";
            xlSheet.Cell(2, 18).Value = "Detail";

            xlSheet.Cell(1, 19).Value = "Inspection";
            xlSheet.Cell(2, 19).Value = "Status";
            xlSheet.Cell(2, 20).Value = "Alarm Type";
            xlSheet.Cell(2, 21).Value = "Collected By";
            xlSheet.Cell(2, 22).Value = "Collected On";
            xlSheet.Cell(2, 23).Value = "Notes";
            xlSheet.Cell(2, 24).Value = "Photo Link";
            xlSheet.Cell(2, 25).Value = "Attachment Link";
            #endregion

            var lastRow = 0;
            for (int i = 0; i < Equipmentlist.Count; i++)
            {
                xlSheet.Cell(i + 3, 1).Value = Equipmentlist[i].equipmentId;
                xlSheet.Cell(i + 3, 2).Value = Equipmentlist[i].equipmentName;
                xlSheet.Cell(i + 3, 3).Value = Equipmentlist[i].equipmentSystem;
                xlSheet.Cell(i + 3, 4).Value = Equipmentlist[i].equipmentType;

                xlSheet.Cell(i + 3, 5).Value = Equipmentlist[i].building;
                xlSheet.Cell(i + 3, 6).Value = Equipmentlist[i].floor;
                xlSheet.Cell(i + 3, 7).Value = Equipmentlist[i].room;
                xlSheet.Cell(i + 3, 8).Value = Equipmentlist[i].zone;

                xlSheet.Cell(i + 3, 9).Value = Equipmentlist[i].docLink;
                xlSheet.Cell(i + 3, 10).Value = Equipmentlist[i].docPhoto;

                if (Equipmentlist[i].docLink != string.Empty)
                { xlSheet.Cell(i + 3, 9).Hyperlink = new XLHyperlink(@Equipmentlist[i].docLink); }
                if (Equipmentlist[i].docPhoto != string.Empty)
                { xlSheet.Cell(i + 3, 10).Hyperlink = new XLHyperlink(@Equipmentlist[i].docPhoto); }

                xlSheet.Cell(i + 3, 11).Value = Equipmentlist[i].classification;
                xlSheet.Cell(i + 3, 12).Value = Equipmentlist[i].materialType;
                xlSheet.Cell(i + 3, 13).Value = Equipmentlist[i].consequencePriority;
                xlSheet.Cell(i + 3, 14).Value = Equipmentlist[i].opeationStatus;

                xlSheet.Cell(i + 3, 15).Value = Equipmentlist[i].manufacturer;
                xlSheet.Cell(i + 3, 16).Value = Equipmentlist[i].year;
                xlSheet.Cell(i + 3, 17).Value = Equipmentlist[i].degradationInfo;
                xlSheet.Cell(i + 3, 18).Value = Equipmentlist[i].detail;

                xlSheet.Cell(i + 3, 19).Value = Equipmentlist[i].inspectionStatus;
                xlSheet.Cell(i + 3, 20).Value = Equipmentlist[i].alarmType;
                xlSheet.Cell(i + 3, 21).Value = Equipmentlist[i].collectedBy;
                xlSheet.Cell(i + 3, 22).Value = Equipmentlist[i].collectedOn;
                xlSheet.Cell(i + 3, 23).Value = Equipmentlist[i].notes;

                xlSheet.Cell(i + 3, 24).Value = Equipmentlist[i].inspectionPhotoLink;
                xlSheet.Cell(i + 3, 25).Value = Equipmentlist[i].attachmentLink;

                if (Equipmentlist[i].inspectionPhotoLink != string.Empty)
                { xlSheet.Cell(i + 3, 24).Hyperlink = new XLHyperlink(@Equipmentlist[i].inspectionPhotoLink); }
                if (Equipmentlist[i].attachmentLink != string.Empty)
                { xlSheet.Cell(i + 3, 25).Hyperlink = new XLHyperlink(@Equipmentlist[i].attachmentLink); }
                lastRow = i + 3;
            }
            ExcelStyle(xlSheet,lastRow);
            var folderPath = SaveFile();
            xlWorkbook.SaveAs(folderPath + "/" + "EquipmentListFromSql"+ ".xlsx");
        }

        private void ExcelStyle(IXLWorksheet xlSheet,int lastRow)
        {
            var range1 = xlSheet.Range("A1:D1");
            var range2 = xlSheet.Range("E1:H1");
            var range3 = xlSheet.Range("I1:L1");
            var range4 = xlSheet.Range("M1:N1");
            var range5 = xlSheet.Range("O1:R1");
            var range6 = xlSheet.Range("S1:Y1");

            range1.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range2.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range3.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range4.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range5.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range6.Merge().Style.Font.SetBold().Font.FontSize = 14;

            var rangeString = $"A1:Y{lastRow}";
            xlSheet.Range(rangeString).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            xlSheet.Range(rangeString).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            xlSheet.Range(rangeString).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            xlSheet.Range(rangeString).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            xlSheet.Range(rangeString).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            xlSheet.Range(rangeString).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        }

        private void FilterBySelection(IDbConnection connection)
        {
            // columnFilter = "equipmentId,building";
            var columnFilterList = new List<string>();
            if (IDCheck) columnFilterList.Add("equipmentId");
            if (NameCheck) columnFilterList.Add("equipmentName");
            if (BarCodeCheck) columnFilterList.Add("equipmentSystem");
            if (TypeCheck) columnFilterList.Add("equipmentType");

            if (BuildingCheck) columnFilterList.Add("building");
            if (LevelCheck) columnFilterList.Add("floor");
            if (RoomCheck) columnFilterList.Add("room");
            if (ZoneCheck) columnFilterList.Add("zone");

            if (DocLinkCheck) columnFilterList.Add("docLink");
            if (PhotoLinkCheck) columnFilterList.Add("docPhoto");

            if (ClassificationCheck) columnFilterList.Add("classification");
            if (MaterialTypeCheck) columnFilterList.Add("materialType");
            if (ConsequencePriorityCheck) columnFilterList.Add("consequencePriority");
            if (OperationStatusCheck) columnFilterList.Add("opeationStatus");

            if (ManufacturerCheck) columnFilterList.Add("manufacturer");
            if (YearCheck) columnFilterList.Add("year");
            if (DegradationInforCheck) columnFilterList.Add("degradationInfo");
            if (DetailCheck) columnFilterList.Add("detail");

            if (InspectionStatusCheck) columnFilterList.Add("inspectionStatus");
            if (YearCheck) columnFilterList.Add("alarmType");
            if (DegradationInforCheck) columnFilterList.Add("collectedBy");
            if (DetailCheck) columnFilterList.Add("collectedOn");
            if (YearCheck) columnFilterList.Add("notes");
            if (DegradationInforCheck) columnFilterList.Add("inspectionPhotoLink");
            if (DetailCheck) columnFilterList.Add("attachmentLink");

            columnFilter = string.Join(",", columnFilterList);
            Equipmentlist = connection.Query<Equipment>($"select {columnFilter} from EquipmentList where {SearchKeyWord} = '{ SerachValue }'").ToList();//direct call Sql
        }
        private void SaveToExcelByFilter()
        {
            var xlWorkbook = new XLWorkbook();
            var xlSheet = xlWorkbook.Worksheets.Add("Equipment List");

            var columnFilterList = columnFilter.Split(',');
            for (int i = 0; i < columnFilterList.Length; i++)
            {
                xlSheet.Cell(1, i+1).Value = columnFilterList[i];             
            }

            for (int i = 0; i < Equipmentlist.Count; i++)
            {
                var columnCount = 0;
                if (IDCheck){ columnCount++;xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].equipmentId;}
                if (NameCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].equipmentName; }
                if (BarCodeCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].equipmentSystem; }
                if (TypeCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].equipmentType; }

                if (BuildingCheck) {columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].building;}
                if (LevelCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].floor; }
                if (RoomCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].room; }
                if (ZoneCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].zone; }

                if (DocLinkCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].docLink; }
                if (PhotoLinkCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].docPhoto; }

                if (ClassificationCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].classification; }
                if (MaterialTypeCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].materialType; }
                if (ConsequencePriorityCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].consequencePriority; }
                if (OperationStatusCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].opeationStatus; }

                if (ManufacturerCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].manufacturer; }
                if (YearCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].year; }
                if (DegradationInforCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].degradationInfo; }
                if (DetailCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].detail; }

                if (InspectionStatusCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].inspectionStatus; }
                if (AlarmTypeCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].alarmType; }
                if (CollectedByCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].collectedBy; }
                if (CollectedOnCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].collectedOn; }
                if (NotesCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].notes; }
                if (InspectPhotoCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].inspectionPhotoLink; }
                if (AttachmentLinkCheck) { columnCount++; xlSheet.Cell(i + 2, columnCount).Value = Equipmentlist[i].attachmentLink; }
            }


            var folderPath = SaveFile();
            xlWorkbook.SaveAs(folderPath + "/" + "EquipmentListFromSql_Selection" + ".xlsx");
        }

    }
}
