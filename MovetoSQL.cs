using System.Data.SqlClient;

namespace ChecklistReview
{
    class MovetoSQL
    {
        public static void action(System.Data.DataTable table)
        {
            string ConnectionString = "";//need to find the connection string
            SqlConnection sqlConn = new SqlConnection(ConnectionString);
            //SqlDataAdapter adapter = new SqlDataAdapter(string.Format("SELECT * FROM {0}", cmboTableOne.SelectedItem), sqlConn);
            //using (new SqlCommandBuilder(adapter))
           // {
           //     try
           //     {
            //        adapter.Fill(table);
            //        sqlConn.Open();
            //        adapter.Update(table);
            //        sqlConn.Close();
             //   }
              //  catch (Exception es)
              //  {
              //      MessageBox.Show(es.Message, @"SQL Connection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
              //  }
          //  }
        }
    }
}
