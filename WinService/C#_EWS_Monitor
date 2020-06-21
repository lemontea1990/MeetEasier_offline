using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;

namespace WindowsService_EWS
{
    public partial class EWS_Monitor : ServiceBase
    {
        public EWS_Monitor()
        {
            InitializeComponent();
        }

        private System.Timers.Timer aTimer;

        protected override void OnStart(string[] args)
        {

            //timer will only define the time between two service 
            // will define another timer in the ews() to stop current timer, once that loop finished then continue

            aTimer = new System.Timers.Timer();           
            aTimer.Interval = 60000 * 2;
            aTimer.Elapsed += new System.Timers.ElapsedEventHandler(EWS);
            
            aTimer.Enabled = true;
            aTimer.AutoReset = true;

            writeLog("Start Service" );


        }

		//write log to file 
        protected void get_time(object source, System.Timers.ElapsedEventArgs e)
        {
            
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter("D:\\" + 1 + "WS_EWS_log.txt", true))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Start.");

            }
        }

		//use threads to arrange data write 
        protected void writeLog(string log)
        {

            ParameterizedThreadStart pth_log = new ParameterizedThreadStart(writeLogWait);
            Thread th_log = new Thread(pth_log);
            th_log.IsBackground = true;
            th_log.Start(log);
            
        }

		//define lock to avlid file write conflict 
        static ReaderWriterLockSlim log_locker = new ReaderWriterLockSlim();

        protected void writeLogWait(object obj)
        {
            string log = obj as string;

            log_locker.EnterWriteLock();

            try
            {

                using (System.IO.StreamWriter sw = new System.IO.StreamWriter("D:\\" + 1 + "WS_EWS_log.txt", true))
                {
                    sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + log);
                }
            }
            catch (Exception)
            { }

            finally {

                log_locker.ExitWriteLock();
            }
        }


		//main EWS entry
        protected void EWS(object source, System.Timers.ElapsedEventArgs e)
        {
            writeLog("Start EWS Operation");


            Stopwatch s_watch = new Stopwatch();
            s_watch.Start();

   
		
			//clean the mid-transfer DB table 
            string str_delete = "delete  from Meeting_Status_Loop";
            SqlCommand sql_delete = new SqlCommand(str_delete, sql_conn());
            sql_delete.ExecuteNonQuery();

            sql_conn().Close();
            
			
            //check the credential in the DB whether it still works or not.
            string str_credential = "select * from credential ";

            DataSet ds = new DataSet();
            ds = db_do(str_credential);
            if (ds.Tables[0].Rows.Count < 1)
            {
                string str_err = "insert into err_msg values ('The credential information was not found','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "')";
                SqlCommand cmd_str_err = new SqlCommand(str_err, sql_conn());

                cmd_str_err.ExecuteNonQuery();
                sql_conn().Close();

            }


            //stopwatch
            s_watch.Stop();
            writeLog(s_watch.ElapsedMilliseconds.ToString());            
            s_watch.Start();


            //start to connect EWS and get the ews data
            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2013);
			
            es.Credentials = new WebCredentials(ds.Tables[0].Rows[0][0].ToString(), ds.Tables[0].Rows[0][1].ToString(), "Your Domian");

            string u_name = ds.Tables[0].Rows[0][0].ToString();
            string pwd = ds.Tables[0].Rows[0][1].ToString();


            //insert your own ews uri
            es.Url = new Uri("Your EWS Uri");

            try { es.GetRoomLists(); }
            catch (Exception et)
            {
                string str_err = "insert into err_msg values ('Please update the credential!!','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "')";
                SqlCommand cmd_str_err = new SqlCommand(str_err, sql_conn());

                cmd_str_err.ExecuteNonQuery();

                sql_conn().Close();

            }

            EmailAddressCollection myRoomLists = es.GetRoomLists();

            //stopwatch
            s_watch.Stop();
            writeLog("the time of get roomlits is:"+s_watch.ElapsedMilliseconds.ToString());

            DataTable dt_sql = new DataTable();
            dt_sql.Columns.Add("RoomName", typeof(string));
            dt_sql.Columns.Add("StartTime", typeof(string));
            dt_sql.Columns.Add("EndTime", typeof(string));
            dt_sql.Columns.Add("Organzier", typeof(string));
            dt_sql.Columns.Add("Location", typeof(string));
            dt_sql.Columns.Add("TimeStamp", typeof(string));
            dt_sql.Columns.Add("Subject", typeof(string));
            dt_sql.Columns.Add("Opt_att", typeof(string));
            dt_sql.Columns.Add("Req_att", typeof(string));

            EmailAddress address = new EmailAddress() ;

            foreach (EmailAddress address_list in myRoomLists)
            {             

                if (address_list.ToString().Contains("SSMR"))
                {
                    address = address_list;
                }
            }

            System.Collections.ObjectModel.Collection<EmailAddress> room_addresses = es.GetRooms(address);

            aTimer.Stop();
            writeLog( "  now we stop the timer");

            int room_count = room_addresses.Count;
            for (int i = 0; i < room_count; i++)
            {
                

                string str_mailbox = room_addresses[i].Address.ToString();
                string str_room_name = room_addresses[i].Name.ToString();

                if (str_mailbox.Contains("internal"))
                {
                    str_mailbox = str_mailbox.Replace("internal.", "");
                }

                
                try
                {
                    Stopwatch sw_dt = new Stopwatch();
                    sw_dt.Start();


                    FolderId fid = new FolderId(WellKnownFolderName.Calendar, str_mailbox);
                    CalendarFolder cfolder = CalendarFolder.Bind(es, fid);

                    CalendarView cview = new CalendarView(DateTime.Now, DateTime.Now.AddDays(1));
                    cview.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Organizer, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.IsMeeting, AppointmentSchema.Location, AppointmentSchema.LegacyFreeBusyStatus, AppointmentSchema.DisplayTo, AppointmentSchema.DisplayCc);


                    FindItemsResults<Appointment> appo = cfolder.FindAppointments(cview);

                    if (appo != null)
                    {      
                        foreach (Appointment appoi in appo.ToList())
                        {
                            string str_sub = "";
                            if (appoi.Subject is null)
                            {
                                str_sub = "";
                            }
                            else
                            {
                                str_sub = appoi.Subject.ToString().Replace("'", " ");
                            }


                            dt_sql.Rows.Add(str_room_name, appoi.Start.ToString("yyyy-MM-dd HH:mm:ss.ffff"), appoi.End.ToString("yyyy-MM-dd HH:mm:ss.ffff"), null_replace(appoi.Organizer.Name),
                                null_replace(appoi.Location), System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff"), str_sub, null_replace(appoi.DisplayCc), null_replace(appoi.DisplayTo));


                        }
                    }
                    sw_dt.Stop();
                    writeLog("Insert room:" + str_room_name + "cost time: " + sw_dt.ElapsedMilliseconds.ToString());

                }
                catch (Exception et)
                {
                    // Create an email message and provide it with connection 
                    // configuration information by using an ExchangeService object named service.
                    //EmailMessage message = new EmailMessage(es);
                    //// Set properties on the email message.
                    //message.Subject = "EWS Monitor Service Error!!";
                    //message.Body = "Dears, there is an error when monitor " + str_mailbox + " the exchange server <br><br><br>  the error is:" + et.ToString();
                    //message.ToRecipients.Add("email address");


                    // Send the email message and save a copy.
                    // This method call results in a CreateItem call to EWS.
                    //message.SendAndSaveCopy();

                }

                if (i == room_count - 1)
                {
                    aTimer.Start();
                    writeLog( aTimer.ToString() + "  now we restart the timer");
                }

            }

            writeLog("all rooms operation finsihed");

            Stopwatch sw_dt_sql = new Stopwatch();
            sw_dt_sql.Start();

            try
            {
                
                SqlBulkCopy sql_dt_insert = new SqlBulkCopy(sql_conn());
                sql_dt_insert.BatchSize = 10000;
                sql_dt_insert.BulkCopyTimeout = 60;

                sql_dt_insert.DestinationTableName = "Meeting_Status_Loop";

                for (int c = 0; c < dt_sql.Columns.Count; c++)
                {
                    sql_dt_insert.ColumnMappings.Add(c, c + 1);
                }
                sql_dt_insert.WriteToServer(dt_sql);

                sql_conn().Close();

                string str_insert = "insert into Meeting_Status_Loop values ('end','end','end','end','end','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "','end','end','end')";
                SqlCommand sql_cmd = new SqlCommand(str_insert, sql_conn());
                sql_cmd.ExecuteNonQuery();
               

                //writeLog("get the final one: No." + i + " of meeting room" + str_room_name);

                dt_sql.Clear();
                sql_conn().Close();

            }
            catch (Exception et)
            {
                writeLog("\r\n" + "error when insert end to DB" + et);
            }

            sw_dt_sql.Stop();
            writeLog("\r\n" + " time cost of dt insert is:" + sw_dt_sql.ElapsedMilliseconds.ToString());

        }

        protected SqlConnection sql_conn()
        {
            
			string sql_conn_str = "Your DB connection string";
            SqlConnection sql_conn = new SqlConnection(sql_conn_str);
            sql_conn.Open();

            return sql_conn;
        }
        protected DataSet db_do(string sql_search)
        {

            SqlDataAdapter da = new SqlDataAdapter(sql_search, sql_conn());
            DataSet ds = new DataSet();
            da.Fill(ds, "dbresult");

            sql_conn().Close();

            return ds;
        }

        protected string null_replace(string str_input)
        {
            string str_output = "";

            if (string.IsNullOrEmpty(str_input) is false)
            {
                str_output = str_input.Replace("'", " ");
            }

            return str_output;
        }

        //use mutilable process
        //protected void GetData_EWS(string username, string pwd, Mailbox mailbox, string RoomName)
        protected void GetData_EWS(object obj)
        {
            //List<string> par_get = obj as List<string>;

            DataTable dt_get = obj as DataTable;

            string username = dt_get.Rows[0][0].ToString();
            string pwd = dt_get.Rows[0][1].ToString();
            string mailbox = dt_get.Rows[0][2].ToString();
            string RoomName = dt_get.Rows[0][3].ToString();

            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2013);
            es.Credentials = new WebCredentials(username,pwd, "Your Domain");

            
            es.Url = new Uri("Your own EWS uri");

            Stopwatch sw = new Stopwatch();
            

            try
            {
                sw.Start();


                FolderId fid = new FolderId(WellKnownFolderName.Calendar, mailbox);
                CalendarFolder cfolder = CalendarFolder.Bind(es, fid);

                CalendarView cview = new CalendarView(DateTime.Now, DateTime.Now.AddDays(1));
                cview.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Organizer, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.IsMeeting, AppointmentSchema.Location, AppointmentSchema.LegacyFreeBusyStatus, AppointmentSchema.DisplayTo, AppointmentSchema.DisplayCc);


                FindItemsResults<Appointment> appo = cfolder.FindAppointments(cview);


                sw.Stop();
                writeLog("time cost of get appointments list for"+ RoomName + "is :" + sw.ElapsedMilliseconds.ToString());
                sw.Restart();


                foreach (Appointment appoi in appo.ToList())
                {
                    
                    string str_sub = "";
                    if (appoi.Subject is null)
                    {
                        str_sub = "";
                    }
                    else
                    {
                        str_sub = appoi.Subject.ToString().Replace("'", " ");
                    }



                    if (string.IsNullOrEmpty(appoi.DisplayTo) is false)
                    {

                    }

                    //MessageBox.Show(appoi.Subject + appoi.Organizer + appoi.Location);
                    string str_insert = "insert into Meeting_Status_Loop values ('" + RoomName + "','" + appoi.Start.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "','" +
                       appoi.End.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "','" + null_replace(appoi.Organizer.Name) + "','" + null_replace(appoi.Location) + "','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "',N'" +
                       (str_sub) + "','" + null_replace(appoi.DisplayCc) + "','" + null_replace(appoi.DisplayTo) + "')";

                    SqlCommand sqlCommand = new SqlCommand(str_insert, sql_conn());

                    try
                    {
                        sqlCommand.ExecuteNonQuery();
                        //MessageBox.Show("loop start");

                        writeLog(" sql operation for ." + mailbox + "was successfully finished");
                        sql_conn().Close();

                        //EmailMessage message = new EmailMessage(service);
                        //// Set properties on the email message.
                        //message.Subject = "EWS Monitor Service Sql Operaction successed";
                        //message.Body = "Dears, finished the tasks for sql operation";
                        //message.ToRecipients.Add("Your email address ");
                        //// Send the email message and save a copy.
                        //// This method call results in a CreateItem call to EWS.
                        //message.SendAndSaveCopy();
                    }
                    catch (Exception et)
                    {
                        // Create an email message and provide it with connection 
                        // configuration information by using an ExchangeService object named service.
                        EmailMessage message = new EmailMessage(es);
                        // Set properties on the email message.
                        message.Subject = "EWS Monitor Service Error!!";
                        message.Body = "Dears, there is an error when monitor " + mailbox + " the exchange server <br><br><br>  the error is:" + et.ToString();
                        message.ToRecipients.Add("Your Email address");
                        

                        // Send the email message and save a copy.
                        // This method call results in a CreateItem call to EWS.
                        message.SendAndSaveCopy();

                    }
                    //}


                }

                sw.Stop();
                writeLog("the time cost of SQL operate for" + RoomName + "is : " + sw.ElapsedMilliseconds.ToString());
            }

            catch (Exception et)
            {
                writeLog("there is an error when operate at"+RoomName+ " the error is:\r\n" +et);
            }

        }

        protected override void OnStop()
        {
            writeLog("Service Stopped!");
        }



        

    }
}
