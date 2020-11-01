
using System.Windows;
using ExcelSqlDataExchange.ViewModel;
using System.Data;
using ExcelSqlDataExchange.Support;
using Dapper;
using System;


namespace ExcelSqlDataExchange
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private QueryViewModel qv = new QueryViewModel();
        private UpdateDataViewModel uv = new UpdateDataViewModel();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void QueryOperation(object sender, RoutedEventArgs e)
        {
            DataContext = qv;
        }

        private void UpdateOperation(object sender, RoutedEventArgs e)
        {
            DataContext = uv;
        }

        private void ResetOperation(object sender, RoutedEventArgs e)
        {
            try
            {
                using (IDbConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVal("AssetDB")))
            {
                connection.Execute("EquipmentList_ClearAll");
            }
            MessageBox.Show("The data has been cleared to the SQL Server!");
            }
            catch (Exception a)
            {
                MessageBox.Show(a.Message);
                throw;
            }
        }
    }
}
