using System.Data;
using System.Data.SqlClient;

namespace TeamNavigator
{
    class DataSetSource
    {
        private DataSet _dataSet;
        public DataSetSource()
        {
            _dataSet = new DataSet();

            // string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["teamsConnectionString"].ConnectionString;
            string connStr = "xx";

            SqlConnection cn = new SqlConnection(connStr);

            SqlDataAdapter rda = new SqlDataAdapter("SELECT ROOMNAME FROM rooms", cn);
            rda.Fill(_dataSet, "Rooms");

            SqlDataAdapter tda = new SqlDataAdapter("SELECT TEAMNAME FROM teams", cn);
            tda.Fill(_dataSet, "Teams");

            SqlDataAdapter mda = new SqlDataAdapter("SELECT * FROM members", cn);
            mda.Fill(_dataSet, "Members");
        }

        public DataTable GetTeamMembers(string teamname)
        {
            DataTable table = _dataSet.Tables["Members"];
            table.DefaultView.RowFilter = $"TEAMNAME1='{teamname}' or TEAMNAME2='{teamname}' or TEAMNAME3='{teamname}'";

            return table;
        }

        public DataTable GetAllMembers()
        {
            DataTable table = _dataSet.Tables["Members"];

            return table;
        }

        public DataTable GetTeams()
        {
            DataTable table = _dataSet.Tables["Teams"];
            return table;
        }

        public DataTable GetRooms()
        {
            DataTable table = _dataSet.Tables["Rooms"];
            return table;
        }
    }
}
