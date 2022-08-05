using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace ProjectAddIn5
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.NewProject += new Microsoft.Office.Interop.MSProject._EProjectApp2_NewProjectEventHandler(Application_NewProject);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        void Application_NewProject(Microsoft.Office.Interop.MSProject.Project pj)
        {

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
                    MSProject.Task newTask = pj.Tasks.Add
                            (reader.GetString(1), missing);
                    newTask.Start = reader.GetDateTime(2);
                    newTask.Finish = reader.GetDateTime(3);
                    newTask.Duration = reader.GetValue(5);
                    if (reader.GetValue(4).ToString() != string.Empty)
                    {
                        int pred = reader.GetInt32(4)+counter;
                        newTask.Predecessors = pred.ToString();
                    }

                }
            
            }
        }
    }
}