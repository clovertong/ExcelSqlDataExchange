using System;
using System.Collections.Generic;
using ExcelSqlDataExchange.Support;
using Dapper;
using ExcelSqlDataExchange.Models;
using System.Data;
using System.Windows;
using System.Collections.ObjectModel;
using ClosedXML.Excel;
using System.Linq;

namespace ExcelSqlDataExchange.ViewModel
{
    public class UpdateDataViewModel : ModelView, IRequireViewIdentification
    {
        public ModelView modelView { get; set; }

        public XLWorkbook xlWorkbook { get; set; }

        List<Equipment> Equipmentlist = new List<Equipment>();

        #region Excel

        private List<string> _sheetNames;
        public List<string> SheetNames
        {
            get { return _sheetNames; }
            set
            {
                _sheetNames = value;
                SetPropertyChanged("SheetNames");
            }
        }

        private string _selctedSheet;
        public string SelectSheet
        {
            get { return _selctedSheet; }
            set
            {
                _selctedSheet = value;
                SetPropertyChanged("SelectSheet");
            }
        }

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

        #region UpdataOption Prop
        private string _updateOption = "Get Data by User Input";
        public string UpdateOption
        {
            get { return _updateOption; }
            set
            {
                _updateOption = value;
                SetPropertyChanged("UpdateOption");
            }
        }

        private string _sheet ;
        public string Sheet
        {
            get { return _sheet; }
            set
            {
                _sheet = value;
                SetPropertyChanged("Sheet");
            }
        }
        #endregion

        #region User Input

        #region General Infor

        private string _id = "id1";
        public string ID
        {
            get { return _id; }
            set
            {
                _id = value;
                SetPropertyChanged("ID");
            }
        }

        private string _name= "Equp1";
        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                SetPropertyChanged("Name");
            }
        }

        private string _type=string.Empty;
        public string Type
        {
            get { return _type; }
            set
            {
                _type = value;
                SetPropertyChanged("Type");
            }
        }

        private string _barCode = string.Empty;
        public string BarCode
        {
            get { return _barCode; }
            set
            {
                _barCode = value;
                SetPropertyChanged("BarCode");
            }
        }
        #endregion

        #region Location
        private string _building = "Building1";
        public string Building
        {
            get { return _building; }
            set
            {
                _building = value;
                SetPropertyChanged("Building");
            }
        }

        private string _level = "Basement1";
        public string Level
        {
            get { return _level; }
            set
            {
                _level = value;
                SetPropertyChanged("Level");
            }
        }

        private string _room = "Room";
        public string Room
        {
            get { return _room; }
            set
            {
                _room = value;
                SetPropertyChanged("Room");
            }
        }

        private string _zone = string.Empty;
        public string Zone
        {
            get { return _zone; }
            set
            {
                _zone = value;
                SetPropertyChanged("Zone");
            }
        }


        #endregion

        #region Documentation
        private string _docLink = string.Empty;
        public string DocLink
        {
            get { return _docLink; }
            set
            {
                _docLink = value;
                SetPropertyChanged("DocLink");
            }
        }

        private string _docPhotoLink = string.Empty;
        public string DocPhotoLink
        {
            get { return _docPhotoLink; }
            set
            {
                _docPhotoLink = value;
                SetPropertyChanged("DocPhotoLink");
            }
        }

        #endregion

        #region Classification
        private string _classification = "Uniclass 2015";
        public string Classification
        {
            get { return _classification; }
            set
            {
                _classification = value;
                SetPropertyChanged("Classification");
            }
        }

        private string _materialType = string.Empty;
        public string MaterialType
        {
            get { return _materialType; }
            set
            {
                _materialType = value;
                SetPropertyChanged("MaterialType");
            }
        }

        private string _consequencePriority = "Medium";
        public string ConsequencePriority
        {
            get { return _consequencePriority; }
            set
            {
                _consequencePriority = value;
                SetPropertyChanged("ConsequencePriority");
            }
        }

        private string _opeationStatus = "Normal";
        public string OpeationStatus
        {
            get { return _opeationStatus; }
            set
            {
                _opeationStatus = value;
                SetPropertyChanged("OpeationStatus");
            }
        }

        #endregion

        #region Manufacturer
        private string _manufacturer = string.Empty;
        public string Manufacturer
        {
            get { return _manufacturer; }
            set
            {
                _manufacturer = value;
                SetPropertyChanged("Manufacturer");
            }
        }

        private DateTime _mYear = DateTime.Today;
        public DateTime MM
        {
            get { return _mYear; }
            set
            {
                _mYear = value;
                SetPropertyChanged("MM");
            }
        }

        private string _degradationInfo = string.Empty;
        public string DegradationInfo
        {
            get { return _degradationInfo; }
            set
            {
                _degradationInfo = value;
                SetPropertyChanged("DegradationInfo");
            }
        }

        private string _detail = string.Empty;
        public string Detail
        {
            get { return _detail; }
            set
            {
                _detail = value;
                SetPropertyChanged("Detail");
            }
        }
        #endregion

        #region Inspection
        private string _status = "Closed";
        public string Status
        {
            get { return _status; }
            set
            {
                _status = value;
                SetPropertyChanged("Status");
            }
        }

        private string _alarmType = string.Empty;
        public string AlarmType
        {
            get { return _alarmType; }
            set
            {
                _alarmType = value;
                SetPropertyChanged("AlarmType");
            }
        }

        private string _collectedBy = string.Empty;
        public string CollectedBy
        {
            get { return _collectedBy; }
            set
            {
                _collectedBy = value;
                SetPropertyChanged("CollectedBy");
            }
        }

        private DateTime _collectedOn=DateTime.Today;
        public DateTime CollectedOn
        {
            get { return _collectedOn; }
            set
            {
                _collectedOn = value;
                SetPropertyChanged("CollectedOn"); 
            }
        }

        private string _notes = string.Empty;
        public string Notes
        {
            get { return _notes; }
            set
            {
                _notes = value;
                SetPropertyChanged("Notes");
            }
        }

        private string _inspectionPhotoLink = string.Empty;
        public string InspectionPhotoLink
        {
            get { return _inspectionPhotoLink; }
            set
            {
                _inspectionPhotoLink = value;
                SetPropertyChanged("InspectionPhotoLink");
            }
        }

        private string _attachmentLink = string.Empty;
        public string AttachmentLink
        {
            get { return _attachmentLink; }
            set
            {
                _attachmentLink = value;
                SetPropertyChanged("AttachmentLink");
            }
        }
        #endregion
        #endregion

        #region DelegateCommond

        private DelegateCommand _import;
        public DelegateCommand Import
        {
            get { return _import; }
            set
            {
                _import = value;
                SetPropertyChanged("Import");
            }
        }
        private DelegateCommand _run;
        public DelegateCommand Run
        {
            get { return _run; }
            set
            {
                _run = value;
                SetPropertyChanged("Run");
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

        public UpdateDataViewModel()
        {
            modelView = new ModelView();
            Import = new DelegateCommand(ImportAction);
            Run = new DelegateCommand(RunAction);
            _viewId = Guid.NewGuid();
        }

        private void ImportAction()
        {
            var excelFile = ImportFile();
            xlWorkbook = new XLWorkbook(excelFile);
            SheetNames = new List<string>();
            foreach (IXLWorksheet worksheet in xlWorkbook.Worksheets)
            {
                SheetNames.Add(worksheet.Name);
            }
        }
        private void RunAction()
        {  
            try
            {
                Equipmentlist.Clear();
                if (UpdateOption == "Get Data by Excel")
            {
                UpdateThroughExcel();
            }
            else {
                UpdateThroughUserInput();
                SheetNames = new List<string>();
                SelectSheet = string.Empty;
            }
            }
            catch (Exception a)
            {
                MessageBox.Show(a.Message);
                WindowManager.CloseWindow(ViewID);
                throw;

            }
        }

        private void UpdateThroughUserInput()
        {
            using (IDbConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVal("AssetDB")))
            {
                var equipment=new Equipment {
                    equipmentId = ID, equipmentName = Name,  equipmentSystem = BarCode, equipmentType = Type,
                    building = Building,floor = Level,room = Room,zone = Zone,
                    docLink = DocLink,docPhoto = DocPhotoLink,
                    classification = Classification,materialType = MaterialType, consequencePriority = ConsequencePriority,opeationStatus = OpeationStatus,
                    manufacturer =Manufacturer,year=MM.ToShortDateString(),degradationInfo=DegradationInfo,detail=Detail,
                    inspectionStatus = Status,alarmType=AlarmType,collectedBy=CollectedBy,collectedOn=CollectedOn.ToShortDateString(),
                    notes=Notes,inspectionPhotoLink=InspectionPhotoLink,attachmentLink=AttachmentLink
                };
                Equipmentlist.Add(equipment);
                connection.Execute(
                 "dbo.EquipmentList_Insert " +
                 "@equipmentId, @equipmentName, @equipmentSystem, @equipmentType," +
                 "@building,@floor,@room,@zone," +
                 "@docLink,@docPhoto," +
                 "@classification,@materialType," +
                 "@consequencePriority,@opeationStatus," +
                 "@manufacturer,@year,@degradationInfo,@detail," +
                 "@inspectionStatus,@alarmType,@collectedBy," +
                 "@collectedOn,@notes,@inspectionPhotoLink,@attachmentLink"
                    , Equipmentlist);

            }
            MessageBox.Show("The data has been uploaded to the SQL Server!");
            Equipmentlist.Clear();
        }

        private void UpdateThroughExcel()
        {
            ReadDataFromExcel();
            using (IDbConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVal("AssetDB")))
            {
                connection.Execute(
                 "dbo.EquipmentList_Insert " +
                 "@equipmentId, @equipmentName, @equipmentSystem, @equipmentType," +
                 "@building,@floor,@room,@zone," +
                 "@docLink,@docPhoto," +
                 "@classification,@materialType," +
                 "@consequencePriority,@opeationStatus," +
                 "@manufacturer,@year,@degradationInfo,@detail," +
                 "@inspectionStatus,@alarmType,@collectedBy," +
                 "@collectedOn,@notes,@inspectionPhotoLink,@attachmentLink"
                 , Equipmentlist);
            }
            LinkToDataGrid();
            MessageBox.Show("The data has been uploaded to the SQL Server!");
            Equipmentlist.Clear();
        }

        private void ReadDataFromExcel()
        {
            IXLWorksheet xlsheet = xlWorkbook.Worksheet(SelectSheet);           
            var rowCount = xlsheet.Column(1).Cells().Count();
            var startRow = 4;
            for (int i = startRow; i <= rowCount; i++)
            {
                string id = xlsheet.Cell(i, 1).Value.ToString();
                string equimentName = xlsheet.Cell(i, 2).Value.ToString();
                string barCode = xlsheet.Cell(i, 3).Value.ToString();
                string type = xlsheet.Cell(i, 4).Value.ToString();

                string building = xlsheet.Cell(i, 5).Value.ToString();
                string level = xlsheet.Cell(i, 6).Value.ToString();
                string room = xlsheet.Cell(i, 7).Value.ToString();
                string zone = xlsheet.Cell(i, 8).Value.ToString();

                string docLink = xlsheet.Cell(i, 9).Value.ToString();
                string docPhoto = xlsheet.Cell(i, 10).Value.ToString();

                string classification = xlsheet.Cell(i, 11).Value.ToString();
                string materialType = xlsheet.Cell(i, 12).Value.ToString();
                string consequencePriority = xlsheet.Cell(i, 13).Value.ToString();
                string opeationStatus = xlsheet.Cell(i, 14).Value.ToString();

                string manufacturer = xlsheet.Cell(i, 15).Value.ToString();
                string year = xlsheet.Cell(i, 16).Value.ToString();
                string degradationInfo = xlsheet.Cell(i, 17).Value.ToString();
                string detail = xlsheet.Cell(i, 18).Value.ToString();

                string inspectionStatus = xlsheet.Cell(i, 19).Value.ToString();
                string alarmType = xlsheet.Cell(i, 20).Value.ToString();
                string collectedBy = xlsheet.Cell(i, 21).Value.ToString();
                string collectedOn = xlsheet.Cell(i, 22).Value.ToString();

                string notes = xlsheet.Cell(i, 23).Value.ToString();
                string inspectionPhotoLink = xlsheet.Cell(i, 24).Value.ToString();
                string attachmentLink = xlsheet.Cell(i, 25).Value.ToString();

                var equipment=new Equipment
                {
                    equipmentId = id,
                    equipmentName = equimentName,
                    equipmentSystem = barCode,
                    equipmentType = type,
                    classification = classification,
                    materialType = materialType,
                    consequencePriority = consequencePriority,
                    opeationStatus = opeationStatus,
                    building = building,
                    floor = level,
                    room = room,
                    zone = zone,
                    docLink = docLink,
                    docPhoto = docPhoto,
                    manufacturer = manufacturer,
                    year = year,
                    degradationInfo = degradationInfo,
                    detail = detail,
                    inspectionStatus = inspectionStatus,
                    alarmType = alarmType,
                    collectedBy = collectedBy,
                    collectedOn = collectedOn,
                    notes = notes,
                    inspectionPhotoLink = inspectionPhotoLink,
                    attachmentLink = attachmentLink
                };

                Equipmentlist.Add(equipment);
            }

        }

        private void LinkToDataGrid()
        {
            DataGrid.Clear();
            foreach (var equipment in Equipmentlist)
                DataGrid.Add(equipment);
        }
    }
}