using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using MySql.Data.MySqlClient;

namespace TourQueryManager
{
    public partial class FrmQueryWorkingPage : Form
    {
     

        static string queryIdWorking = null;
        static string mysqlConnStr = Properties.Settings.Default.mysqlConnStr;
        MySqlConnection frmQueryWorkingMysqlConn = new MySqlConnection(mysqlConnStr);
        MySqlTransaction frmQueryWorkingMysqlTransaction = null;
        MySqlDataAdapter frmQueryWorkingMysqlDataAdaptor = new MySqlDataAdapter();
        DataSet frmQueryWorkingDataSet = null;
        DataSet dataSetHotelCity = new DataSet();
        DataSet dataSetHotelRating = new DataSet();
        DataSet dataSetHotelName = new DataSet();
        DataSet dataSetRoomType = new DataSet();
        DataSet dataSetHotelMealPlan = new DataSet();
        MySqlCommand command = new MySqlCommand();
        String columnNameForMealPlan = "";
        int hotelPricePerDay = 0;

        public FrmQueryWorkingPage(string queryId)
        {
            InitializeComponent();
            queryIdWorking = queryId;
        }

        private void FrmQueryWorkingPage_Load(object sender, EventArgs e)
        {
            /* data to be initialized when loading page */
            Text = "Working on QUERYID (" + queryIdWorking + ")";
            
           
            try
            {
                frmQueryWorkingMysqlConn.Open();
            }
            catch (Exception erropen)
            {
                MessageBox.Show("connection cannot be opened because " + erropen.Message + "");
                Close();
            }

           
           
            frmQueryWorkingDataSet = new DataSet();
            string mysqlSelectQueryStr = "SELECT * FROM `queries` WHERE `queryid` = " + queryIdWorking + "";
            command.Connection = frmQueryWorkingMysqlConn;
            command.CommandText = mysqlSelectQueryStr;
            try
            {
                frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
               
                frmQueryWorkingMysqlDataAdaptor.Fill(frmQueryWorkingDataSet, "QUERYID_DATA");
            }
            catch (Exception errquery)
            {
                MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
            }
            string querydetails = "";
            querydetails = querydetails + "QUERY START DATE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["querystartdate"].ToString();
            querydetails = querydetails + "\r\nPLACE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["place"].ToString();
            querydetails = querydetails + "\r\nDESTINATION COVERED = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["destinationcovered"].ToString();
            querydetails = querydetails + "\r\nFROM DATE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["fromdate"].ToString();
            querydetails = querydetails + "\r\nTO DATE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["todate"].ToString();
            querydetails = querydetails + "\r\nADULTS = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["adults"].ToString();
            querydetails = querydetails + "\r\nCHILDREN = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["children"].ToString();
            querydetails = querydetails + "\r\nBABIES = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["babies"].ToString();
            querydetails = querydetails + "\r\nROOM COUNT = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["roomcount"].ToString();
            querydetails = querydetails + "\r\nMEAL = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["meal"].ToString();
            querydetails = querydetails + "\r\nHOTEL CATEGORY = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["hotelcategory"].ToString();
            querydetails = querydetails + "\r\nARRIVAL DATE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["arrivaldate"].ToString();
            querydetails = querydetails + "\r\nDEPARTURE DATE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["departuredate"].ToString();
            querydetails = querydetails + "\r\nARRIVAL CITY = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["arrivalcity"].ToString();
            querydetails = querydetails + "\r\nDEPARTURE CITY = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["departurecity"].ToString();
            querydetails = querydetails + "\r\nVEHICAL CATEGORY = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["vehicalcategory"].ToString();
            querydetails = querydetails + "\r\nREQUIREMENT = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["requirement"].ToString();
            querydetails = querydetails + "\r\nBUDGET = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["budget"].ToString();
            querydetails = querydetails + "\r\nNOTE = " + frmQueryWorkingDataSet.Tables["QUERYID_DATA"].Rows[0]["note"].ToString();
            querydetails = querydetails + "";
            txtboxQueryDetails.Text = querydetails;

            //day no info 
            numericUpDownWorkingDayNo.Value = 1;
            numericUpDownNoOfCars.Value = 0;
            CmbboxWrkngCarType.Enabled = false;
            CmbboxWrkngCarPurpose.Enabled = false;

            //arrivalDate
            dateTimePickerWorkingArrivalDate.Enabled = true;


            //Sector Info
            string mysqlSelectSectorStr = "SELECT DISTINCT hotelarea FROM `hotelinfo` ";
            command.CommandText = mysqlSelectSectorStr;
            try
            {
                frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
              
                frmQueryWorkingMysqlDataAdaptor.Fill(frmQueryWorkingDataSet, "SECTOR_DATA");
                
            }
            catch (Exception errquery)
            {
                MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
            }

            DataRow dr = frmQueryWorkingDataSet.Tables["SECTOR_DATA"].NewRow();
            dr[0] = "SELECT SECTOR";

            frmQueryWorkingDataSet.Tables["SECTOR_DATA"].Rows.InsertAt(dr, 0);

            CmbboxWrkngHtlSector.Items.Clear();
            //  CmbboxWrkngHtlSector.Items.Add("SELECT SECTOR");
            CmbboxWrkngHtlSector.ValueMember = "hotelarea";
            CmbboxWrkngHtlSector.DisplayMember = "hotelarea";
            CmbboxWrkngHtlSector.DataSource = frmQueryWorkingDataSet.Tables["SECTOR_DATA"];
            CmbboxWrkngHtlSector.SelectedIndex = 0;
            CmbboxWrkngSeasonType.SelectedIndex = 0;
           // CmbboxWrkngHtlSector.SelectedValue = 0;



           
        }

        
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception errReleaseObj)
            {
                Debug.WriteLine("Error in releasing object :: " + errReleaseObj.Message);
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void FrmQueryWorkingPage_FormClosing(object sender, FormClosingEventArgs e)
        {
           
            
        }

        private void CmbboxWrkngHtlSector_SelectedIndexChanged(object sender, EventArgs e)
        {
          


            if (CmbboxWrkngHtlSector.SelectedIndex == 0)
            {
                CmbboxWrkngHtlLocation.Text = "--- Select sector first ---";
                return;
            }
            else
            {
                dataSetHotelCity.Reset();
                CmbboxWrkngHtlLocation.DataSource = null;



                command.CommandText = "SELECT DISTINCT `hotelcity` FROM `hotelinfo`  WHERE `hotelarea` = '" + CmbboxWrkngHtlSector.Text + "' ORDER BY `hotelcity`";
                // MessageBox.Show("Query  " + command.CommandText);
                try
                {
                    frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
                    // frmQueryWorkingMysqlDataAdaptor = new MySqlDataAdapter(mysqlSelectSectorStr, frmQueryWorkingMysqlConn);
                    frmQueryWorkingMysqlDataAdaptor.Fill(dataSetHotelCity, "HOTEL_CITY_DATA");

                    DataRow dr = dataSetHotelCity.Tables["HOTEL_CITY_DATA"].NewRow();
                    dr[0] = "SELECT CITY";

                    dataSetHotelCity.Tables["HOTEL_CITY_DATA"].Rows.InsertAt(dr, 0);


                    // CmbboxWrkngHtlLocation.Items.Add("SELECT LOCATION");
                    CmbboxWrkngHtlLocation.ValueMember = "hotelcity";
                    CmbboxWrkngHtlLocation.DisplayMember = "hotelcity";
                    CmbboxWrkngHtlLocation.DataSource = dataSetHotelCity.Tables["HOTEL_CITY_DATA"];
                    // CmbboxWrkngHtlLocation.SelectedValue = 0;
                    CmbboxWrkngHtlLocation.SelectedIndex = 0;

                }
                catch (Exception errquery)
                {
                    MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
                }
            }
           
        }

        private void CmbboxWrkngHtlLocation_SelectedIndexChanged(object sender, EventArgs e)
        {



            if (CmbboxWrkngHtlLocation.SelectedIndex == 0)
            {
                CmbboxWrkngHtlHotelRating.Text = "--- Select city first ---";
                return;
            }
            else
            {

                CmbboxWrkngHtlHotelRating.DataSource = null;
                dataSetHotelRating.Reset();


                command.CommandText = "SELECT DISTINCT `hotelrating` FROM `hotelinfo`  WHERE `hotelarea` = '" + CmbboxWrkngHtlSector.Text + "' AND `hotelcity` = '" + CmbboxWrkngHtlLocation.Text + "'  ORDER BY `hotelrating`";
                // MessageBox.Show("Query  " + command.CommandText);
                try
                {
                    frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
                    // frmQueryWorkingMysqlDataAdaptor = new MySqlDataAdapter(mysqlSelectSectorStr, frmQueryWorkingMysqlConn);
                    frmQueryWorkingMysqlDataAdaptor.Fill(dataSetHotelRating, "HOTEL_RATING");

                    DataRow dr = dataSetHotelRating.Tables["HOTEL_RATING"].NewRow();
                    dr[0] = "SELECT HOTEL RATING";

                    dataSetHotelRating.Tables["HOTEL_RATING"].Rows.InsertAt(dr, 0);



                    CmbboxWrkngHtlHotelRating.ValueMember = "hotelrating";
                    CmbboxWrkngHtlHotelRating.DisplayMember = "hotelrating";
                    CmbboxWrkngHtlHotelRating.DataSource = dataSetHotelRating.Tables["HOTEL_RATING"];
                    //CmbboxWrkngHtlHotelRating.SelectedValue = 0;
                    CmbboxWrkngHtlHotelRating.SelectedIndex = 0;

                }
                catch (Exception errquery)
                {
                    MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
                }

            }



        }

        private void CmbboxWrkngHtlHotel_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (CmbboxWrkngHtlHotel.SelectedIndex > 0 )
            {
                CmbboxWrkngHtlRoomType.DataSource = null;
                dataSetRoomType.Reset();
                CmbboxWrkngHtlRoomType.Items.Clear();

                // numericUpDownWorkingRoomNo.Value = 1;

                /* 
                            command.CommandText = "SELECT `idhotelinfo` FROM `hotelinfo`  WHERE `hotelarea` = '" + CmbboxWrkngHtlSector.Text + "' AND `hotelcity` = '" + CmbboxWrkngHtlLocation.Text + "' AND `hotelrating` = '" + CmbboxWrkngHtlHotelRating.Text + "' AND `hotelname` = '" + CmbboxWrkngHtlHotel.SelectedValue + "'  LIMIT 1";
                             MessageBox.Show("Query::  " + command.CommandText);

                           MySqlDataReader reader = command.ExecuteReader();

                            try
                            {
                                if (reader.Read())
                                {
                                    a = reader.GetInt32(0);
                                    MessageBox.Show("idhotelnfo " + a);

                                }
                            }
                            catch (Exception errquery)
                            {
                                MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
                            }
                            */
                string seasontype = CmbboxWrkngSeasonType.Text;
                string seasonyear = dateTimePickerWorkingArrivalDate.Value.ToString("yyyy");
               // MessageBox.Show(seasontype +"  and  "+ seasonyear);
                command.CommandText = "SELECT  `idhotelrates` , `roomtype` FROM  `hotelrates`  WHERE  `idhotelinfo` = " + CmbboxWrkngHtlHotel.SelectedValue + " AND  `seasontype` = '" + seasontype + "' AND  `seasonyear` = '" + seasonyear + "' ORDER BY `idhotelrates`";
                // MessageBox.Show("manoj  " + command.CommandText);
                try
                {
                    frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
                    // frmQueryWorkingMysqlDataAdaptor = new MySqlDataAdapter(mysqlSelectSectorStr, frmQueryWorkingMysqlConn);
                    frmQueryWorkingMysqlDataAdaptor.Fill(dataSetRoomType, "ROOM_TYPE");

                    DataRow dr = dataSetRoomType.Tables["ROOM_TYPE"].NewRow();
                    dr[0] = 0;
                    dr[1] = "SELECT ROOM TYPE";
                    dataSetRoomType.Tables["ROOM_TYPE"].Rows.InsertAt(dr, 0);


                    CmbboxWrkngHtlRoomType.ValueMember = "idhotelrates";
                    CmbboxWrkngHtlRoomType.DisplayMember = "roomtype";
                    CmbboxWrkngHtlRoomType.DataSource = dataSetRoomType.Tables["ROOM_TYPE"];

                    CmbboxWrkngHtlRoomType.SelectedIndex = 0;

                }
                catch (Exception errquery)
                {
                    MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
                }

            }
            else
            {
                CmbboxWrkngHtlRoomType.Text = "--- Select hotelName first ---";
                return;
               

            }




        }

        private void buttonHotelAddRow_Click(object sender, EventArgs e)
        {
            


            int index = dataGrdVuWorkingHotels.Rows.Add();
            
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNo"].Value = numericUpDownWorkingDayNo.Value.ToString();
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayDate"].Value = dateTimePickerWorkingArrivalDate.Value.ToString("yyyy-MM-dd");
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayArrivalTime"].Value = dttmpkrWorkingArvlTime.Value.ToString("hh:mm:ss tt");
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayArea"].Value = CmbboxWrkngHtlSector.Text;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayCity"].Value = CmbboxWrkngHtlLocation.Text;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayName"].Value = CmbboxWrkngHtlHotel.Text; ;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNoOfRooms"].Value = (numericUpDownWorkingRoomNo.Value -1).ToString();
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDaySim"].Value = chkBoxWorkingSim.Checked.ToString();
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayGuide"].Value = chkBoxWorkingGuide.Checked.ToString();
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayAdditionalCost"].Value = txtboxWorkingAdditionalCost.Text;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayPrice"].Value = hotelPricePerDay.ToString();
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNarrationHdr"].Value = txtboxWorkingNarrationHeader.Text;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNarration"].Value = txtboxWorkingNarration.Text;
            dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayUserComment"].Value = txtboxWorkingUserComment.Text;
            numericUpDownWorkingDayNo.Value = numericUpDownWorkingDayNo.Value + 1;
            hotelPricePerDay = 0;
            dateTimePickerWorkingArrivalDate.Visible = false;
            dateTimePickerWorkingArrivalDate.Value = dateTimePickerWorkingArrivalDate.Value.AddDays(1);
            numericUpDownWorkingRoomNo.Value = 1;
            btnWorkingAddRoom.Enabled = false;
            numericUpDownNoOfCars.Value = 0;
            chkBoxWorkingSim.Checked = false;
            chkBoxWorkingGuide.Checked = false;

            //reset all fields
            txtboxWorkingAdditionalCost.Text = "0";
            txtboxWorkingFlightCost.Text = "0";

            CmbboxWrkngHtlHotel.SelectedIndex = 0;
            CmbboxWrkngHtlRoomType.SelectedIndex = 0;
            
            CmbboxWrkngHtlLocation.SelectedIndex = 0;
            CmbboxWrkngHtlSector.SelectedIndex = 0;
           
            
            

            CmbboxWrkngHtlLocation.DataSource=null;
            CmbboxWrkngHtlHotel.DataSource = null;
            CmbboxWrkngHtlRoomType.DataSource = null;

        }

        private void ButtonWorkingDone_Click(object sender, EventArgs e)
        {
            /* upload data of work done in the database */
            int workingHotelDayRowsCount = dataGrdVuWorkingHotels.RowCount;
            int workingHotelRoomsRowsCount = dataGridViewRoomsInfo.RowCount;
            if (workingHotelDayRowsCount < 1 && workingHotelRoomsRowsCount < 1)
            {
                MessageBox.Show("Working area is empty.");
                return;
            }
            else { 
            bool querySuccessDay = true;
            bool querySuccessRoom = true;
            bool querySuccess = true;
                MySqlCommand mySqlCommandButtonWorkingDone = frmQueryWorkingMysqlConn.CreateCommand();
            frmQueryWorkingMysqlTransaction = frmQueryWorkingMysqlConn.BeginTransaction();
            mySqlCommandButtonWorkingDone.Connection = frmQueryWorkingMysqlConn;
            mySqlCommandButtonWorkingDone.Transaction = frmQueryWorkingMysqlTransaction;
            mySqlCommandButtonWorkingDone.CommandType = CommandType.Text;
            string mysqlInsertQueryStr = "INSERT INTO `queryworkingday` ( " +
                "`queryid`, " +
                "`dayno`, " +
                "`area`, " +
                "`city`, " +
                "`hotel`, " +
                "`pricehotel`, " +
                "`narrationhdr`," +
                "`narration`, " +
                "`additionalcost`, " +
                "`usercomment` , " +
                "`sim`, " +
                "`guide` ) " +
                "VALUES ( " +
                "@queryid_var, " +
                "@dayno_var, " +
                
                "@area_var, " +
                "@city_var, " +
                "@hotel_var, " +
                "@pricehotel_var, " +
                "@narrationhdr_var, " +
                "@narration_var, " +
                "@additionalcost_var, " +
                "@usercomment_var, " +
                "@sim_var, " +
                "@guide_var )";
            mySqlCommandButtonWorkingDone.CommandText = mysqlInsertQueryStr;
            mySqlCommandButtonWorkingDone.Prepare();
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@queryid_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@dayno_var", 1);
            
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@area_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@city_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@hotel_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@pricehotel_var", 1);
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@narrationhdr_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@narration_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@additionalcost_var", 1);
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@usercomment_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@sim_var", "Text");
            mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@guide_var", "Text");

                for (int index = 0; index < workingHotelDayRowsCount; index++)
            {
                /* insert row data in database */
                mySqlCommandButtonWorkingDone.Parameters["@queryid_var"].Value = queryIdWorking;
                mySqlCommandButtonWorkingDone.Parameters["@dayno_var"].Value = Convert.ToInt32(dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNo"].Value);
                
                mySqlCommandButtonWorkingDone.Parameters["@area_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayArea"].Value.ToString();
                mySqlCommandButtonWorkingDone.Parameters["@city_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayCity"].Value.ToString();
                mySqlCommandButtonWorkingDone.Parameters["@hotel_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayName"].Value.ToString();
                mySqlCommandButtonWorkingDone.Parameters["@pricehotel_var"].Value = Convert.ToInt32(dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayPrice"].Value);
                    mySqlCommandButtonWorkingDone.Parameters["@narrationhdr_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNarrationHdr"].Value.ToString(); 
                    mySqlCommandButtonWorkingDone.Parameters["@narration_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayNarration"].Value.ToString();
                    mySqlCommandButtonWorkingDone.Parameters["@additionalcost_var"].Value = Convert.ToInt32(dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayAdditionalCost"].Value);
                mySqlCommandButtonWorkingDone.Parameters["@usercomment_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayUserComment"].Value.ToString();
                    mySqlCommandButtonWorkingDone.Parameters["@sim_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDaySim"].Value.ToString();
                mySqlCommandButtonWorkingDone.Parameters["@guide_var"].Value = dataGrdVuWorkingHotels.Rows[index].Cells["wrkHotelDayGuide"].Value.ToString();
                    try
                {
                    mySqlCommandButtonWorkingDone.ExecuteNonQuery();
                    
                }
                catch (Exception errquery)
                {
                    MessageBox.Show("queryworkingday  : Error while executing insert query because:\n" + errquery.Message);
                    querySuccessDay = false;
                    break;
                }

                    }


                // room data

                string mysqlInsertRoomDetailQueryStr = "INSERT INTO `queryworkingroom` ( " +
                "`idqueryworkingday`, " +
                "`queryid`, " +
                "`dayno`, " +
                "`roomtype`, " +
                "`mealplan`, " +
                "`extrabed`, " +
                "`roomprice` ) " +
                "VALUES ( " +
                "@idqueryworkingday_var, " +
                "@queryid_room_var, " +
                "@dayno_room_var, " +

                "@roomtype_var, " +
                "@mealplan_var, " +
                "@extrabed_var, " +
                
                "@roomprice_var )";
                mySqlCommandButtonWorkingDone.CommandText = mysqlInsertRoomDetailQueryStr;
                mySqlCommandButtonWorkingDone.Prepare();
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@idqueryworkingday_var", 1);
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@queryid_room_var", "Text");
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@dayno_room_var", 1);

                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@roomtype_var", "Text");
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@mealplan_var", "Text");
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@extrabed_var", "Text");
                mySqlCommandButtonWorkingDone.Parameters.AddWithValue("@roomprice_var", 1);

                for (int index = 0; index < workingHotelRoomsRowsCount; index++)
                {
                    /* insert row data in database */
                    command.CommandText = "SELECT `idqueryworkingday` FROM `queryworkingday` WHERE `queryid` ="+ queryIdWorking + " AND `dayno` ="+ Convert.ToInt32(dataGridViewRoomsInfo.Rows[index].Cells["wrkDayNo"].Value);
                    int idQueryWorkingDay =(Int32) command.ExecuteScalar();
                    MessageBox.Show(command.CommandText+"  and  "+ idQueryWorkingDay);
                    mySqlCommandButtonWorkingDone.Parameters["@idqueryworkingday_var"].Value = idQueryWorkingDay;
                    mySqlCommandButtonWorkingDone.Parameters["@queryid_room_var"].Value = queryIdWorking;
                    mySqlCommandButtonWorkingDone.Parameters["@dayno_room_var"].Value = Convert.ToInt32(dataGridViewRoomsInfo.Rows[index].Cells["wrkDayNo"].Value);

                    mySqlCommandButtonWorkingDone.Parameters["@roomtype_var"].Value = dataGridViewRoomsInfo.Rows[index].Cells["wrkRoomType"].Value.ToString();
                    mySqlCommandButtonWorkingDone.Parameters["@mealplan_var"].Value = dataGridViewRoomsInfo.Rows[index].Cells["wrkMealPlan"].Value.ToString();
                    mySqlCommandButtonWorkingDone.Parameters["@extrabed_var"].Value = dataGridViewRoomsInfo.Rows[index].Cells["wrkExtraBed"].Value.ToString();
                    mySqlCommandButtonWorkingDone.Parameters["@roomprice_var"].Value = Convert.ToInt32(dataGridViewRoomsInfo.Rows[index].Cells["wrkPrice"].Value);
                    
                    try
                    {
                        mySqlCommandButtonWorkingDone.ExecuteNonQuery();

                    }
                    catch (Exception errquery)
                    {
                        MessageBox.Show("queryworkingRoom :: Error while executing insert query because:\n" + errquery.Message);
                        querySuccessRoom = false;
                        break;
                    }
                }
                    //end room data


                    if (querySuccessDay && querySuccessRoom)
            {
                MessageBox.Show("Query Executed successfully now changing query status");
                mysqlInsertQueryStr = "UPDATE `queries` SET " +
                    "`querylastupdatetime` = NOW(), " +
                    "`querycurrentstate` = " + Properties.Resources.queryStageDoneByUser + " " +
                    "WHERE " +
                    "queryid = " + queryIdWorking;
                mySqlCommandButtonWorkingDone.CommandText = mysqlInsertQueryStr;
                mySqlCommandButtonWorkingDone.Prepare();
                try
                {
                    mySqlCommandButtonWorkingDone.ExecuteNonQuery();

                }
                catch (Exception errquery)
                {
                    MessageBox.Show("Error while executing insert query because:\n" + errquery.Message);
                    querySuccess = false;
                }
            }
            try
            {
                if (querySuccess)
                {
                    frmQueryWorkingMysqlTransaction.Commit();
                }
                else
                {
                    frmQueryWorkingMysqlTransaction.Rollback();
                }
            }
            catch (Exception errcommit)
            {
                MessageBox.Show("Error while executing insert query because:\n" + errcommit.Message);
            }
            MessageBox.Show("Work for Query id = " + queryIdWorking + " is done successfully ");
            Close();
                }
        }

        private void ButtonWorkingCancel_Click(object sender, EventArgs e)
        {
            /* exit without doing anything*/
            Close();
        }

        private void numericUpDownRoomNo_ValueChanged(object sender, EventArgs e)
        {

        }

        private void lblWorkingNarration_Click(object sender, EventArgs e)
        {

        }

        private void lblWorkingHotelMealPlan_Click(object sender, EventArgs e)
        {

        }

        private void CmbboxWrkngHtlHotelRating_SelectedIndexChanged(object sender, EventArgs e)
        {



            if (CmbboxWrkngHtlHotelRating.SelectedIndex == 0)
            {
                CmbboxWrkngHtlHotel.Text = "--- Select rating first ---";
                return;
            }
            else
            {
                CmbboxWrkngHtlHotel.DataSource = null;
                dataSetHotelName.Reset();
                CmbboxWrkngHtlHotel.Items.Clear();



                command.CommandText = "SELECT DISTINCT `idhotelinfo`, `hotelname` FROM `hotelinfo`  WHERE `hotelarea` = '" + CmbboxWrkngHtlSector.Text + "' AND `hotelcity` = '" + CmbboxWrkngHtlLocation.Text + "' AND `hotelrating` = '" + CmbboxWrkngHtlHotelRating.Text + "'  ORDER BY `hotelname`";
                // MessageBox.Show("Query  " + command.CommandText);
                try
                {
                    frmQueryWorkingMysqlDataAdaptor.SelectCommand = command;
                    // frmQueryWorkingMysqlDataAdaptor = new MySqlDataAdapter(mysqlSelectSectorStr, frmQueryWorkingMysqlConn);
                    frmQueryWorkingMysqlDataAdaptor.Fill(dataSetHotelName, "HOTEL_NAME");

                    DataRow dr = dataSetHotelName.Tables["HOTEL_NAME"].NewRow();
                    dr[0] = 0;
                    dr[1] = "SELECT HOTEL ";

                    dataSetHotelName.Tables["HOTEL_NAME"].Rows.InsertAt(dr, 0);


                    CmbboxWrkngHtlHotel.ValueMember = "idhotelinfo";

                    CmbboxWrkngHtlHotel.DisplayMember = "hotelname";
                    CmbboxWrkngHtlHotel.DataSource = dataSetHotelName.Tables["HOTEL_NAME"];
                    // CmbboxWrkngHtlHotel.SelectedValue = 0;
                    CmbboxWrkngHtlHotel.SelectedIndex = 0;

                }
                catch (Exception errquery)
                {
                    MessageBox.Show("Query cannot be executed because " + errquery.Message + "");
                }
            }
        }

        private void CmbboxWrkngHtlRoomType_SelectedIndexChanged(object sender, EventArgs e)
        {
            CmbboxWrkngHtlMealPlan.Items.Clear();
           CmbboxWrkngHtlMealPlan.Items.AddRange(new String[] { "SELECT MEAL TYPE", "EPAI", "CPAI", "MAPAI", "APAI"});
            CmbboxWrkngHtlMealPlan.SelectedIndex = 0;
        }

        private void numericUpDownExtraBed_ValueChanged(object sender, EventArgs e)
        {/*
            if (numericUpDownNoOfPersons.Value != 0)
            {
                CmbboxWrkngHtlMealPlan.Items.Clear();
                CmbboxWrkngHtlMealPlan.Items.AddRange(new String[] { "SELECT MEAL TYPE", "SINGLE EPAI", "SINGLE CPAI", "SINGLE MAPAI", "SINGLE APAI", "DOUBLE EPAI", "DOUBLE CPAI", "DOUBLE MAPAI", "DOUBLE APAI", "EXT_BED EPAI", "EXT_BED CPAI", "EXT_BED MAPAI", "EXT_BED APAI" });
                CmbboxWrkngHtlMealPlan.SelectedIndex = 0;
            }*/
        }

        private void CmbboxWrkngHtlMealPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
           columnNameForMealPlan = "";
            if (CmbboxWrkngHtlMealPlan.SelectedIndex == 0)
            { btnWorkingAddRoom.Enabled = false; }
            else
            {
                btnWorkingAddRoom.Enabled = true;
                int noOfPersons = (Int32)numericUpDownNoOfPersons.Value;

                switch (noOfPersons)
                {
                    case 1:
                        {
                            if (CmbboxWrkngHtlMealPlan.Text.Equals("EPAI"))
                            {
                                columnNameForMealPlan = "mealepaipricesingle";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("CPAI"))
                            {
                                columnNameForMealPlan = "mealcpaipricesingle";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("MAPAI"))
                            {
                                columnNameForMealPlan = "mealmapaipricesingle";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("APAI"))
                            {
                                columnNameForMealPlan = "mealapaipricesingle";
                            }
                            break;  
                        }
                    case 2:
                        {
                            if (CmbboxWrkngHtlMealPlan.Text.Equals("EPAI"))
                            {
                                columnNameForMealPlan = "mealepaipricedouble";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("CPAI"))
                            {
                                columnNameForMealPlan = "mealcpaipricedouble";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("MAPAI"))
                            {
                                columnNameForMealPlan = "mealmapaipricedouble";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("APAI"))
                            {
                                columnNameForMealPlan = "mealapaipricedouble";
                            }
                            break;
                        }
                    case 3:
                        {
                            if (CmbboxWrkngHtlMealPlan.Text.Equals("EPAI"))
                            {
                                columnNameForMealPlan = " (mealepaipricedouble + mealepaipriceextbed) AS Total ";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("CPAI"))
                            {
                                columnNameForMealPlan = " (mealcpaipricedouble + mealcpaipriceextbed) AS Total ";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("MAPAI"))
                            {
                                columnNameForMealPlan = " (mealmapaipricedouble + mealmapaipriceextbed) AS Total ";
                            }
                            else if (CmbboxWrkngHtlMealPlan.Text.Equals("APAI"))
                            {
                                columnNameForMealPlan = " (mealapaipricedouble + mealapaipriceextbed) AS Total ";
                            }
                            break;
                        }

                    default: break;
                }
               
               


            }
        }

        private void btnWorkingAddRoom_Click(object sender, EventArgs e)
        {
            

            if (columnNameForMealPlan.Equals("") && ((Int32)CmbboxWrkngHtlRoomType.SelectedValue <= 0))
            {
                MessageBox.Show("PLEASE ENTER ALL INFORMATION");
            }
            else {

                int roomNo = (Int32)numericUpDownWorkingRoomNo.Value;
                
                command.CommandText = "SELECT  " + columnNameForMealPlan + "  FROM  `hotelrates` WHERE `idhotelrates` = " + CmbboxWrkngHtlRoomType.SelectedValue;
               
                // int price = (Int32)command.ExecuteScalar();
                int price = 0;
                
                using (MySqlDataReader reader = command.ExecuteReader())
                {
                    while(reader.Read())
                    price = Convert.ToInt32(reader[0]);
                }
                int rows = dataGridViewRoomsInfo.Rows.Count;
               int index = dataGridViewRoomsInfo.Rows.Add();
                dataGridViewRoomsInfo.Rows[index].Cells["wrkDayNo"].Value = numericUpDownWorkingDayNo.Value.ToString();
                dataGridViewRoomsInfo.Rows[index].Cells["wrkRoom"].Value = numericUpDownWorkingRoomNo.Value.ToString();
                dataGridViewRoomsInfo.Rows[index].Cells["wrkRoomType"].Value = CmbboxWrkngHtlRoomType.Text;
                dataGridViewRoomsInfo.Rows[index].Cells["wrkNoOfPersons"].Value = numericUpDownNoOfPersons.Value.ToString();
                dataGridViewRoomsInfo.Rows[index].Cells["wrkMealPlan"].Value = CmbboxWrkngHtlMealPlan.Text;                
                dataGridViewRoomsInfo.Rows[index].Cells["wrkExtraBed"].Value = (numericUpDownNoOfPersons.Value == 3 ? 1 : 0) ;
                dataGridViewRoomsInfo.Rows[index].Cells["WrkPrice"].Value = price.ToString();
                numericUpDownWorkingRoomNo.Value = roomNo + 1;
                hotelPricePerDay += price;
            }


            btnWorkingAddRoom.Enabled = false;
            numericUpDownNoOfPersons.Value = 1;

        }

        private void numericUpDownNoOfCars_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownNoOfCars.Value == 0)
            { CmbboxWrkngCarType.Enabled = false;
                CmbboxWrkngCarPurpose.Enabled = false;
            }
            else
            {
                CmbboxWrkngCarType.Enabled = true;
                CmbboxWrkngCarPurpose.Enabled = true;
            }
        }

        private void dateTimePickerWorkingArrivalDate_ValueChanged(object sender, EventArgs e)
        {

        }

       
    }
}
