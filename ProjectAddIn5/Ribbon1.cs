using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ProjectAddIn5
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void Show_All_Tasks_Click(object sender, RibbonControlEventArgs e)
        {
            var pj = Globals.ThisAddIn.Application.ActiveProject;
            string connection = "Server=DESKTOP-DPGMQGD;Database=MyTasks;Trusted_Connection=True";
            string query1 = "SELECT * FROM dbo.Task_1";
            string query2 = "SELECT * FROM dbo.Task_2";


            using (SqlConnection con = new SqlConnection(connection))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query1, con);
                SqlDataReader reader = cmd.ExecuteReader();
                int counter = 0;
                int counter1 = 0;
                while (reader.Read())
                {
                    counter1++;
                    TaskCreation();
                }
                cmd.Dispose();
                reader.Close();
                cmd = new SqlCommand(query2, con);
                reader = cmd.ExecuteReader();
                counter = counter1;
                while (reader.Read())
                {
                    TaskCreation();
                }
                reader.Close();
                cmd.Dispose();
                con.Close();

                void TaskCreation()
                {
                    Microsoft.Office.Interop.MSProject.Task newTask = pj.Tasks.Add
                            (reader.GetString(1));
                    newTask.Start = reader.GetDateTime(2);
                    newTask.Finish = reader.GetDateTime(3);
                    newTask.Duration = reader.GetValue(5);
                    if (reader.GetValue(4).ToString() != string.Empty)
                    {
                        int pred = reader.GetInt32(4) + counter;
                        newTask.Predecessors = pred.ToString();
                    }

                }

            }

        }

    }
}
