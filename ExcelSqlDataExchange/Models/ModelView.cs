using System;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;


namespace ExcelSqlDataExchange.Models
{
    public class ModelView : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        //[NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void SetPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _outputfilePath;


        public string OutputFilePath
        {
            get { return _outputfilePath; }
            set
            {
                _outputfilePath = value;
                OnPropertyChanged(nameof(OutputFilePath));
            }
        }

        public string FilePath { get; set; }

        public string ImportFile()
        {
            // save your current directory  
            string currentDirectory = Environment.CurrentDirectory;
            using (OpenFileDialog thisDialog = new OpenFileDialog())
            {
                thisDialog.RestoreDirectory = false;
                thisDialog.Filter = "XLSX File(*.xlsx)|*.xlsx|XLSM File(*.xlsm)|*.xlsm|All Files (*.*)|*.*";

                thisDialog.Multiselect = false;
                thisDialog.FilterIndex = 1;


                if (thisDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedDirectory = Environment.CurrentDirectory;  // OpenFileDialog changed this value.   
                    Environment.CurrentDirectory = currentDirectory; // reset the property with the first value.  
                    thisDialog.InitialDirectory = selectedDirectory;// by doing this, it will open the last closed folder
                }
                return thisDialog.FileName;
            }
        }

        public string ExportFile()
        {
            // save your current directory  
            string currentDirectory = Environment.CurrentDirectory;
            using (OpenFileDialog thisDialog = new OpenFileDialog())
            {
                thisDialog.RestoreDirectory = false;
                thisDialog.Filter = "XLSX File(*.xlsx)|*.xlsx|All Files (*.*)|*.*";

                thisDialog.Multiselect = false;
                thisDialog.FilterIndex = 1;


                if (thisDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedDirectory = Environment.CurrentDirectory;  // OpenFileDialog changed this value.   
                    Environment.CurrentDirectory = currentDirectory; // reset the property with the first value.  
                    thisDialog.InitialDirectory = selectedDirectory;// by doing this, it will open the last closed folder
                }
                // MessageBox.Show(thisDialog.FileName);
                return thisDialog.FileName;
            }
        }

        public string SaveFile()
        {
            string filePath = null;
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            string selectedDirectory = Environment.CurrentDirectory;

            dialog.InitialDirectory = selectedDirectory;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                filePath = dialog.FileName;
            }
            return filePath;
        }
    }


}

