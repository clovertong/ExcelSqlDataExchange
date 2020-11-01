
namespace ExcelSqlDataExchange
{
    public class Equipment
    {
        public int id { get; set; }
        #region  General
        public string equipmentId { get; set; }
        public string equimentName { get; set; }
        public string equipmentType { get; set; }
        public string barCode { get; set; }

        #endregion

        #region  Location
        public string building { get; set; }
        public string floor { get; set; }
        public string room { get; set; }
        public string zone { get; set; }
        #endregion

        #region  Documentation
        public string docLink { get; set; }
        public string docPhoto { get; set; }
        #endregion

        #region  Classification
        public string classification { get; set; }
        public string materialType { get; set; }
        public string consequencePriority { get; set; }//high low
        public string opeationStatus { get; set; }
        #endregion

        #region  Manufacturer
        public string manufacturer { get; set; }
        public string year { get; set; }
        public string degradationInfo { get; set; }
        public string detail { get; set; }
        #endregion

        #region  Inspection
        public string inspectionStatus { get; set; }
        public string alarmType { get; set; }
        public string collectedBy { get; set; }
        public string collectedOn { get; set; }
        public string notes { get; set; }
        public string inspectionPhotoLink { get; set; }
        public string attachmentLink { get; set; }
        #endregion

    }
}
