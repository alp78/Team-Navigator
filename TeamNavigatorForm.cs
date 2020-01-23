using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TeamNavigator
{
    public partial class TeamNavigatorForm : Form
    {
        Assembly myAssembly;
        
        LyncClient _client = null;
        Contact _contact = null;
        ContactSubscription _contactSubscription;
        string sip;
        string username;
        string teamname;
        string activity;
        string seatcode;
        RoundButton[] vacantseats;
        Control[] foundButton;
        RoundButton button;
        DataTable membersTable = new DataTable();
        List<Member> membersList = new List<Member>();
        List<Contact> contactsList = new List<Contact>();
        DataTable teamsTable = new DataTable();
        BindingSource _teamsBindingSource = new BindingSource();
        BindingSource _roomsBindingSource = new BindingSource();
        BindingSource _membersBindingSource = new BindingSource();
        DataSetSource _dataSetSource = new DataSetSource();
        DataTable fullMembersTable = new DataTable();
        DataTable roomsTable = new DataTable();
        List<ContactInformationType>  _infoTypesList = new List<ContactInformationType>();
        List<Member> dummyList = new List<Member> { new Member("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "") };
        int greenCounter = 0;
        int redCounter = 0;
        int yellowCounter = 0;
        int greyCounter = 0;
        int blackCounter = 0;
        string currentColor;

        //private WaitForm _waitForm;

        //protected void ShowWaitForm()
        //{
        //    // don't display more than one wait form at a time
        //    if (_waitForm != null && !_waitForm.IsDisposed)
        //    {
        //        return;
        //    }

        //    _waitForm = new WaitForm();
        //    _waitForm.TopMost = true;
        //    _waitForm.StartPosition = FormStartPosition.CenterScreen;
        //    _waitForm.Show();
        //    _waitForm.Refresh();

        //    // force the wait window to display for at least 700ms so it doesn't just flash on the screen
        //    System.Threading.Thread.Sleep(200);
        //    Application.Idle += OnLoaded;
        //}

        //private void OnLoaded(object sender, EventArgs e)
        //{
        //    Application.Idle -= OnLoaded;
        //    _waitForm.Close();
        //}

        public TeamNavigatorForm()
        {
            StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            TeamsComboBox.SelectedIndexChanged += new EventHandler(TeamsComboBox_SelectedIndexChanged);
        }
        private void TeamNavigatorForm_Load(object sender, EventArgs e)
        {
            // hide vacant seats
            vacantseats = new RoundButton[] {DW4, DW18, DW19, SD1, SD2, SD4, SD6, SD8,
            SD15, SD18, SD34, ID1, ID2, ID7, ID8, HR7, FS16, FS23, BI7, FI1, FI5, FI17,
            FI18, FI22, FI26, FI27, FI30, GO11, GO26, GO27, GO29, GO30, RS6, RS7, RS11, TP1, TP2, RE3};

            foreach (RoundButton button in vacantseats)
            {
                button.BackColor = Color.DarkGray;
                button.Visible = false;
            }

            string[] teams = new string[3];

            myAssembly = Assembly.GetExecutingAssembly();

            // load default profile pic
            using (Stream photoStream = myAssembly.GetManifestResourceStream("TeamNavigator.Resources.default_profile_pic.jpg") as Stream)
            {
                if (photoStream != null)
                {
                    Bitmap bm = new Bitmap(photoStream);
                    ContactPhotoBox.Image = bm;
                }
            }

            //ShowWaitForm();

            _client = LyncClient.GetClient();

            _infoTypesList.Add(ContactInformationType.Activity);
            _contactSubscription = _client.ContactManager.CreateSubscription();

            roomsTable = _dataSetSource.GetRooms();
            teamsTable = _dataSetSource.GetTeams();
            fullMembersTable = _dataSetSource.GetAllMembers();

            // create Lync Contact list + initialize presence on buttons + initialize counters
            foreach (DataRow row in fullMembersTable.Rows)
            {
                // create contacts list and subscribe to presence
                sip = row["Sip"].ToString();
                _contact = _client.ContactManager.GetContactByUri(sip);
                if(_contact != null)
                {
                    contactsList.Add(_contact);
                    _contactSubscription.AddContact(_contact);
                    _contactSubscription.Subscribe(ContactSubscriptionRefreshRate.High, _infoTypesList);
                    _contact.ContactInformationChanged += new EventHandler<ContactInformationChangedEventArgs>(PeerContact_ContactInformationChanged);
                }

                // initialize buttons + counters
                seatcode = row["Seatcode"].ToString();
                if (!string.IsNullOrEmpty(seatcode))
                {
                    foundButton = Controls.Find(seatcode, true);

                    if (foundButton != null && foundButton.Length > 0)
                    {
                        button = foundButton[0] as RoundButton;
                    }
                }
                activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();
                switch (activity)
                {
                    case "Available":
                        if (button != null)
                        {
                            button.BackColor = Color.LimeGreen;
                            greenCounter++;
                        }
                        break;
                    case "Busy":
                        if (button != null)
                        {
                            button.BackColor = Color.Red;
                            redCounter++;
                        }
                        break;
                    case "Do not disturb":
                        if (button != null)
                        {
                            button.BackColor = Color.Red;
                            redCounter++;
                        }
                        break;
                    case "In a meeting":
                        if (button != null)
                        {
                            button.BackColor = Color.Red;
                            redCounter++;
                        }
                        break;
                    case "In a call":
                        if (button != null)
                        {
                            button.BackColor = Color.Red;
                            redCounter++;
                        }
                        break;
                    case "Be right back":
                        if (button != null)
                        {
                            button.BackColor = Color.Gold;
                            yellowCounter++;
                        }
                        break;
                    case "Off work":
                        if (button != null)
                        {
                            button.BackColor = Color.Gold;
                            yellowCounter++;
                        }
                        break;
                    case "Away":
                        if (button != null)
                        {
                            button.BackColor = Color.Gold;
                            yellowCounter++;
                        }
                        break;
                    default:
                        if (button != null)
                        {
                            button.BackColor = Color.LightGray;
                            greyCounter++;
                        }
                        break;
                }
            }

            blackCounter = greenCounter + redCounter + yellowCounter + greyCounter;

            // Display initial counters
            ActiveNumberLabel.Text = greenCounter.ToString();
            BusyNumberLabel.Text = redCounter.ToString();
            AwayNumberLabel.Text = yellowCounter.ToString();
            InactiveNumberLabel.Text = greyCounter.ToString();
            TotalNumberLabel.Text = blackCounter.ToString();

            // Rooms Search Box
            DataTable tempRooms;
            RoomsComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            RoomsComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;

            DataView dtviewRooms = new DataView(roomsTable);
            dtviewRooms.Sort = "RoomName ASC";
            tempRooms = dtviewRooms.ToTable();

            // add empty row at beginning
            DataRow emptyRowRooms = tempRooms.NewRow();
            emptyRowRooms[0] = "";
            tempRooms.Rows.InsertAt(emptyRowRooms, 0);

            // bind rooms to combobox
            _roomsBindingSource.DataSource = tempRooms;
            RoomsComboBox.DataSource = _roomsBindingSource;
            RoomsComboBox.DisplayMember = "RoomName";
            RoomsComboBox.ValueMember = "RoomName";


            // Teams Search Box
            DataTable tempTeams;
            TeamsComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            TeamsComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;

            DataView dtviewTeams = new DataView(teamsTable);
            dtviewTeams.Sort = "TeamName ASC";
            tempTeams = dtviewTeams.ToTable();

            // add empty row at beginning
            DataRow emptyRowTeams = tempTeams.NewRow();
            emptyRowTeams[0] = "";
            tempTeams.Rows.InsertAt(emptyRowTeams, 0);


            // bind teams to combobox
            _teamsBindingSource.DataSource = tempTeams;
            TeamsComboBox.DataSource = _teamsBindingSource;
            TeamsComboBox.DisplayMember = "TeamName";
            TeamsComboBox.ValueMember = "TeamName";


            // Members Search Box
            DataTable tempMembers;
            SearchComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            SearchComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;

            DataView dtview = new DataView(fullMembersTable);
            dtview.Sort = "Fullname ASC";
            tempMembers = dtview.ToTable();

            // add empty row at beginning
            DataRow emptyRowMembers = tempMembers.NewRow();
            emptyRowMembers[0] = 0;
            emptyRowMembers[1] = "";
            emptyRowMembers[2] = "";
            tempMembers.Rows.InsertAt(emptyRowMembers, 0);

            SearchComboBox.DataSource = tempMembers;
            SearchComboBox.ValueMember = "Username";
            SearchComboBox.DisplayMember = "Fullname";

            // bind textboxes with members properties
            FullNameTextBox.DataBindings.Add("Text", _membersBindingSource, "FullName");
            EmailTextBox.DataBindings.Add("Text", _membersBindingSource, "Email");
            TitleTextBox.DataBindings.Add("Text", _membersBindingSource, "Title");
            DepartmentTextBox.DataBindings.Add("Text", _membersBindingSource, "Department");
            SeatCodeTextBox.DataBindings.Add("Text", _membersBindingSource, "SeatCode");

            MembersDataGridView.DataSource = _membersBindingSource;

            // hide columns in DataGridView
            MembersDataGridView.Columns["Sip"].Visible = false;
            MembersDataGridView.Columns["IsDirector"].Visible = false;


            SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);


            //if (TeamsComboBox.Items.Count > 0)
            //{
            //    TeamsComboBox.SelectedIndex = 0;
            //}

        }
        private void TeamsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MembersListBox.SelectedIndexChanged -= new EventHandler(MembersListBox_SelectedIndexChanged);

            if (TabPanel.SelectedTab.Text == "Grid")
            {
                MembersDataGridView.Rows.Clear();
                MembersDataGridView.Refresh();
            }

            // bind members to listbox
            if (TeamsComboBox.SelectedIndex != 0)
            {
                
                teamname = TeamsComboBox.SelectedValue.ToString();
                DataTable membersTableTeamChange = _dataSetSource.GetTeamMembers(teamname);

                List<Member> membersListTeamChange = membersTableTeamChange.AsEnumerable().Select(dataRow => new Member(
                    dataRow.Field<string>("Username"), 
                    dataRow.Field<string>("FullName"), 
                    dataRow.Field<string>("Sip"), 
                    dataRow.Field<string>("Email"), 
                    dataRow.Field<string>("Phone"), 
                    dataRow.Field<string>("Mobile"), 
                    dataRow.Field<string>("Manager"), 
                    dataRow.Field<string>("Department"), 
                    dataRow.Field<string>("Title"), 
                    dataRow.Field<string>("TeamName1"), 
                    dataRow.Field<string>("TeamName2"), 
                    dataRow.Field<string>("TeamName3"), 
                    dataRow.Field<string>("SeatCode"), 
                    dataRow.Field<string>("IsTeamLead"), 
                    dataRow.Field<string>("IsManager"), 
                    dataRow.Field<string>("IsOwner"),
                    dataRow.Field<string>("IsDirector")))
                    .Where(x => x.TeamName1 == teamname || x.TeamName2 == teamname || x.TeamName3 == teamname)
                    .OrderBy(x => x.Username)
                    .OrderBy(x => x.IsTeamLead == false)
                    .OrderBy(x => x.IsOwner == false)
                    .OrderBy(x => x.IsManager == false)
                    .OrderBy(x => x.IsDirector == false)
                    .ToList();

                _membersBindingSource.DataSource = membersListTeamChange;
                MembersListBox.DataSource = _membersBindingSource;

                MembersListBox.SelectedIndex = 0;
                MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);

                MembersListBox.ValueMember = "Username";
                MembersListBox.DisplayMember = "Username";

                //if (MembersListBox.Items.Count > 0)
                //{
                //    foreach (var user in membersListTeamChange)
                //    {
                //        if (user.IsOwner == true)
                //        {
                //            MembersListBox.SelectedIndex = MembersListBox.Items.IndexOf(user.Username);
                //            MembersListBox.SelectedValue = user.Username;
                //            break;
                //        }
                //        else if (user.IsTeamLead == true)
                //        {
                //            MembersListBox.SelectedIndex = MembersListBox.Items.IndexOf(user.Username);
                //            MembersListBox.SelectedValue = user.Username;
                //            break;
                //        }
                //        else if (user.IsManager == true)
                //        {
                //            MembersListBox.SelectedIndex = MembersListBox.Items.IndexOf(user.Username);
                //            MembersListBox.SelectedValue = user.Username;
                //            break;
                //        }
                //        else
                //        {
                //            MembersListBox.SelectedIndex = 0;
                //        }
                //    }
                //}
                


                // populate Activity field in MembersDataGridView
                BeginInvoke(new Action(() =>
                {
                    for (int i = 0; i < MembersListBox.Items.Count; i++)
                    {
                        sip = MembersDataGridView.Rows[i].Cells["Sip"].Value.ToString();
                        //username = MembersDataGridView.Rows[i].Cells["Username"].Value.ToString();

                        _contact = _client.ContactManager.GetContactByUri(sip);

                        activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();

                        switch (activity)
                        {
                            case "Available":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.LimeGreen;
                                break;
                            case "Busy":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                break;
                            case "Do not disturb":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                break;
                            case "In a call":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                break;
                            case "In a meeting":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                break;
                            case "Be right back":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                break;
                            case "Off work":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                break;
                            case "Away":
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                break;
                            default:
                                MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Black;
                                break;
                        }
                        MembersDataGridView.Rows[i].Cells["Activity"].Value = activity;
                    }
                }));

                ResetAreas();

                SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);
                SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            }
            else
            {
               
                SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);

                SearchComboBox.SelectedIndex = -1;
                FullNameTextBox.Clear();
                TitleTextBox.Clear();
                SeatCodeTextBox.Clear();
                EmailTextBox.Clear();
                DepartmentTextBox.Clear();
                TeamTextBox.Clear();
                ManagerCheckBox.Checked = false;
                ServiceOwnerCheckBox.Checked = false;
                TeamLeaderCheckBox.Checked = false;
                    
                // load default profile pic
                using (Stream photoStream = myAssembly.GetManifestResourceStream("TeamNavigator.Resources.default_profile_pic.jpg") as Stream)
                {
                    if (photoStream != null)
                    {
                        Bitmap bm = new Bitmap(photoStream);
                        ContactPhotoBox.Image = bm;
                    }
                }
                _membersBindingSource.DataSource = dummyList;

                ResetAreas();

                // initialize buttons color
                foreach (DataRow row in fullMembersTable.Rows)
                {
                    sip = row["Sip"].ToString();
                    _contact = _client.ContactManager.GetContactByUri(sip);

                    // initialize buttons
                    seatcode = row["Seatcode"].ToString();
                    if (!string.IsNullOrEmpty(seatcode))
                    {
                        foundButton = Controls.Find(seatcode, true);

                        if (foundButton != null && foundButton.Length > 0)
                        {
                            button = foundButton[0] as RoundButton;
                        }
                    }
                    activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();

                    if (_contact != null)
                    {
                        switch (activity)
                        {
                            case "Available":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.LimeGreen;
                                    }));
                                }
                                break;
                            case "Busy":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));
                                }
                                break;
                            case "Do not disturb":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));
                                }
                                break;
                            case "In a meeting":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));
                                }
                                break;
                            case "In a call":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));
                                }
                                break;
                            case "Be right back":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));
                                }
                                break;
                            case "Off work":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));
                                }
                                break;
                            case "Away":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));
                                }
                                break;
                            default:
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.LightGray;
                                    }));
                                }
                                break;
                        }
                    }
                    
                }
                MembersListBox.ClearSelected();
                ActivityTextBox.Clear();
                

                using (Stream photoStream = myAssembly.GetManifestResourceStream("TeamNavigator.Resources.default_profile_pic.jpg") as Stream)
                {
                    if (photoStream != null)
                    {
                        Bitmap bm = new Bitmap(photoStream);
                        ContactPhotoBox.Image = bm;
                    }
                }
                //ContactPhotoBox.Refresh();

                BeginInvoke(new Action(() =>
                {
                    ContactBackPhotoBox.BackColor = Color.LightGray;
                }));
                SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            }
        }

        //delegate void StartBlinkProcess(RoundButton blinkButton);
        //delegate void ShowButtonBlink(RoundButton blinkButton);
        private void MembersListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MembersListBox.SelectedIndex != -1)
            {
                string[] teamsArray = new string[3];
                // get Lync contact
                foreach (DataRow row in fullMembersTable.Rows)
                {
                    if (MembersListBox.SelectedValue.ToString() == row["Username"].ToString())
                    {
                        TeamLeaderCheckBox.Checked = (row["IsTeamLead"].ToString() == "Y") ? true : false;
                        ManagerCheckBox.Checked = (row["IsManager"].ToString() == "Y") ? true : false;
                        ServiceOwnerCheckBox.Checked = (row["IsOwner"].ToString() == "Y") ? true : false;
                        sip = row["Sip"].ToString();

                        for (int i = 1; i < 4; i++)
                        {
                            if (row[$"TeamName{i}"] != null)
                            {
                                teamsArray[i - 1] = row[$"TeamName{i}"].ToString();
                            }
                        }
                    }
                }

                // fill TeamTextBox
                TeamTextBox.Text = $"{teamsArray[0]}\r\n{teamsArray[1]}\r\n{teamsArray[2]}";

                // GOC Management global mapping
                if (TeamsComboBox.SelectedValue != null)
                {
                    if (TeamsComboBox.SelectedValue.ToString() == "GOC MANAGEMENT" || TeamsComboBox.SelectedValue.ToString() == "GOC STAFF")
                    {
                        ResetAreas();
                    }
                }

                _contact = _client.ContactManager.GetContactByUri(sip);

                // get contact photo
                try
                {
                    using (Stream photoStream = _contact.GetContactInformation(ContactInformationType.Photo) as Stream)
                    {
                        if (photoStream != null)
                        {
                            Bitmap bm = new Bitmap(photoStream);
                            ContactPhotoBox.Image = bm;
                        }
                    }
                }
                catch (Exception)
                {
                    // load default profile pic
                    using (Stream photoStream = myAssembly.GetManifestResourceStream("TeamNavigator.Resources.default_profile_pic.jpg") as Stream)
                    {
                        if (photoStream != null)
                        {
                            Bitmap bm = new Bitmap(photoStream);
                            ContactPhotoBox.Image = bm;
                        }
                    }
                }

                // get and display activity
                activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();
                switch (activity)
                {
                    case "Available":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.LimeGreen;
                        }));
                        break;
                    case "Busy":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Red;
                        }));
                        break;
                    case "Do not disturb":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Red;
                        }));
                        break;
                    case "In a meeting":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Red;
                        }));
                        break;
                    case "In a call":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Red;
                        }));
                        break;
                    case "Be right back":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Gold;
                        }));
                        break;
                    case "Off work":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Gold;
                        }));
                        break;
                    case "Away":
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.Gold;
                        }));
                        break;
                    default:
                        BeginInvoke(new Action(() =>
                        {
                            ContactBackPhotoBox.BackColor = Color.LightGray;
                        }));
                        break;
                }

                ActivityTextBox.Text = activity;

                // highlight currently selected button
                seatcode = SeatCodeTextBox.Text;
                if (!string.IsNullOrEmpty(seatcode))
                {
                    foundButton = Controls.Find(seatcode, true);
                    RoundButton buttonBlink = foundButton[0] as RoundButton;

                    //StartBlinkProcess startBlink = new StartBlinkProcess(BlinkButton);
                    //startBlink.BeginInvoke(buttonBlink,null,null);
                  
                    //// reset all buttons color
                    foreach (DataRow row in fullMembersTable.Rows)
                    {
                        button = null;
                        seatcode = row["Seatcode"].ToString();
                        sip = row["Sip"].ToString();
                        if (!string.IsNullOrEmpty(seatcode))
                        {
                            foundButton = Controls.Find(seatcode, true);

                            if (foundButton != null && foundButton.Length > 0)
                            {
                                button = foundButton[0] as RoundButton;
                            }
                        }
                        _contact = _client.ContactManager.GetContactByUri(sip);
                        activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();
                        switch (activity)
                        {
                            case "Available":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.LimeGreen;
                                    }));
                                }
                                break;
                            case "Busy":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));
                                }
                                break;
                            case "Do not disturb":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));

                                }
                                break;
                            case "In a meeting":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));

                                }
                                break;
                            case "In a call":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Red;
                                    }));

                                }
                                break;
                            case "Be right back":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));
                                }
                                break;
                            case "Off work":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));

                                }
                                break;
                            case "Away":
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.Gold;
                                    }));

                                }
                                break;
                            default:
                                if (button != null)
                                {
                                    Invoke(new Action(() =>
                                    {
                                        button.BackColor = Color.LightGray;
                                    }));
                                }
                                break;
                        }
                    }

                    BeginInvoke(new Action(() =>
                    {
                        buttonBlink.BackColor = Color.Aqua;
                    }));
                }
            }
        }
        //private void BlinkButton(RoundButton blinkButton)
        //{
        //    if (blinkButton.InvokeRequired == true)
        //    {
        //        ShowButtonBlink showBlink = new ShowButtonBlink(BlinkButton);
        //        BeginInvoke(showBlink, new object[] { blinkButton });
        //    }
        //    else
        //    {
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.White;
        //        Thread.Sleep(50);
        //        button.BackColor = Color.Aqua;
        //    }
        //}
        private void RoomsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            RoomAquariumPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Aquarium" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomCasualPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Casual" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomYellowPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Yellow" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomCornerPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Corner" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomLargePictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Large" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomGreenPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Green" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomBluePictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Blue" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomOneItPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "OneIT" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomVioletPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Violet" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            Room2facesPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "2faces" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RoomCappuccinoPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Cappuccino" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            GreyPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Grey" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            OrangePictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Orange" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;
            RedPictureBox.BackColor = RoomsComboBox.SelectedValue.ToString() == "Red" ? Color.FromArgb(80, 168, 0, 0) : Color.Transparent;

            SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);
            SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
        }
        private void TabPanel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TabPanel.SelectedTab.Text == "Grid" && TeamsComboBox.SelectedIndex !=0 && TeamsComboBox.SelectedIndex != -1)
            {
                // populate Activity field in MembersDataGridView

                for (int i = 0; i < MembersListBox.Items.Count; i++)
                {
                    sip = MembersDataGridView.Rows[i].Cells["Sip"].Value.ToString();
                    username = MembersDataGridView.Rows[i].Cells["Username"].Value.ToString();

                    _contact = _client.ContactManager.GetContactByUri(sip);

                    activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();

                    switch (activity)
                    {
                        case "Available":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.LimeGreen;
                            break;
                        case "Busy":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                            break;
                        case "Do not disturb":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                            break;
                        case "In a call":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                            break;
                        case "In a meeting":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                            break;
                        case "Be right back":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                            break;
                        case "Off work":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                            break;
                        case "Away":
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                            break;
                        default:
                            MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Black;
                            break;
                    }
                    MembersDataGridView.Rows[i].Cells["Activity"].Value = activity;
                }
            }
            SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);
            SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);

        }
        private void PeerContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            if (e.ChangedContactInformation.Contains(ContactInformationType.Activity))
            {
                string activity = null;
                activity = ((Contact)sender).GetContactInformation(ContactInformationType.Activity).ToString();

                string email = ((Contact)sender).Uri.Substring(4).ToLower();

                // update activity text and color if user is selected, and log event 
                if (activity != null)
                {
                    if (EmailTextBox.Text == email)
                    {
                        BeginInvoke(new Action(() => {
                            ActivityTextBox.Text = activity;
                        }));
                    }

                    button = null;
                    bool found = false;
                    int rowIndex = 0;
                    string seatcode = "";
                    Control[] foundButton = null;

                    // Get sender's button
                    foreach (DataRow row in fullMembersTable.Rows)
                    {
                        if (row["Email"].ToString() == email)
                        {
                            seatcode = row["Seatcode"].ToString();
                            if (!string.IsNullOrEmpty(seatcode))
                            {
                                foundButton = Controls.Find(seatcode, true);
                                if (foundButton != null && foundButton.Length > 0)
                                {
                                    button = foundButton[0] as RoundButton;
                                    break;
                                }
                            }

                        }
                    }

                    // change sender's button color and photo background color if selected + update counters
                    switch (activity)
                    {
                        case "Available":
                            if (button != null)
                            {
                                greenCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    
                                    button.BackColor = Color.LimeGreen;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.LimeGreen;
                                    }
                                        
                                }));

                            }
                            break;
                        case "Busy":
                            if (button != null)
                            {
                                redCounter++;
                                currentColor = button.BackColor.ToString();

                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Red;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Red;
                                    }
                                }));
                            }
                            break;
                        case "Do not disturb":
                            if (button != null)
                            {
                                redCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Red;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Red;
                                    }
                                }));

                            }
                            break;
                        case "In a meeting":
                            if (button != null)
                            {
                                redCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Red;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Red;
                                    }

                                }));

                            }
                            break;
                        case "In a call":
                            if (button != null)
                            {
                                redCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Red;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Red;
                                    }

                                }));

                            }
                            break;
                        case "Be right back":
                            if (button != null)
                            {
                                yellowCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Gold;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Gold;
                                    }

                                }));

                            }
                            break;
                        case "Off work":
                            if (button != null)
                            {
                                yellowCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Gold;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Gold;
                                    }

                                }));

                            }
                            break;
                        case "Away":
                            if (button != null)
                            {
                                yellowCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [LightGray]":
                                            greyCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.Gold;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.Gold;
                                    }

                                }));

                            }
                            break;
                        default:
                            if (button != null)
                            {
                                greyCounter++;
                                currentColor = button.BackColor.ToString();
                                Invoke(new Action(() =>
                                {
                                    switch (currentColor)
                                    {
                                        case "Color [LimeGreen]":
                                            greenCounter--;
                                            break;
                                        case "Color [Gold]":
                                            yellowCounter--;
                                            break;
                                        case "Color [Red]":
                                            redCounter--;
                                            break;
                                        default:
                                            break;
                                    }
                                    button.BackColor = Color.LightGray;
                                    if (EmailTextBox.Text == email)
                                    {
                                        ContactBackPhotoBox.BackColor = Color.LightGray;
                                    }

                                }));
                            }
                            break;
                    }

                    // Display updated counters
                    BeginInvoke(new Action(() =>
                    {
                        ActiveNumberLabel.Text = greenCounter.ToString();
                        BusyNumberLabel.Text = redCounter.ToString();
                        AwayNumberLabel.Text = yellowCounter.ToString();
                        InactiveNumberLabel.Text = greyCounter.ToString();
                    }));

                    // update activity cell in DataGridView if current panel is on Grid
                    BeginInvoke(new Action(() =>
                    {
                        if (TabPanel.SelectedTab.Text == "Grid")
                        {
                            // Get index of DataGridView
                            foreach (DataGridViewRow row in MembersDataGridView.Rows)
                            {
                                if (row.Cells["Email"].Value != null && row.Cells["Email"].Value.ToString() == email)
                                {
                                    found = true;
                                    rowIndex = row.Index;
                                    break;
                                }
                            }

                            if (found)
                            {
                                switch (activity)
                                {
                                    case "Available":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.LimeGreen;
                                        break;
                                    case "Busy":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Red;
                                        break;
                                    case "Do not disturb":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Red;
                                        break;
                                    case "In a meeting":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Red;
                                        break;
                                    case "In a call":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Red;
                                        break;
                                    case "Be right back":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Gold;
                                        break;
                                    case "Off work":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Gold;
                                        break;
                                    case "Away":
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Gold;
                                        break;
                                    default:
                                        MembersDataGridView.Rows[rowIndex].Cells["Activity"].Style.ForeColor = Color.Black;
                                        break;
                                }

                                    MembersDataGridView.Rows[rowIndex].Cells["Activity"].Value = activity;
                            }
                        }
                    }));

                    //// Add entry to log
                    //BeginInvoke(new Action(() =>
                    //{
                    //    foreach (DataRow row in fullMembersTable.Rows)
                    //    {
                    //        if (row["Email"].ToString() == email)
                    //        {
                    //            username = row["Username"].ToString();
                    //            teamname = row[$"TeamName1"].ToString();


                    //            if ((!LogTextBox.Text.Contains($"{DateTime.Now.ToString()} - {username} -  {teamname} - {activity}") && (!LogTextBox.Text.Contains($"{DateTime.Now.AddSeconds(-1).ToString()} - {username} -  {teamname} - {activity}"))))
                    //            {
                    //                LogTextBox.Text = $"{DateTime.Now.ToString()} - {username} -  {teamname} - {activity}" + " \r\n" + LogTextBox.Text;
                    //            }
                    //            break;
                    //        }
                    //    }
                    //}));
                }
            }
        }
        private void SearchComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);
            TeamsComboBox.SelectedIndexChanged -= new EventHandler(TeamsComboBox_SelectedIndexChanged);
            MembersListBox.SelectedIndexChanged -= new EventHandler(MembersListBox_SelectedIndexChanged);

            //if (TeamsComboBox.SelectedIndex == -1)
            //{
            //    SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            //}

            if (TabPanel.SelectedTab.Text == "Grid")
            {
                MembersDataGridView.Rows.Clear();
                MembersDataGridView.Refresh();
            }

            if (SearchComboBox.SelectedIndex != 0 && SearchComboBox.SelectedIndex != -1)
            {
                string teamname = "";
                string email = "";
                string username = SearchComboBox.SelectedValue.ToString();
                foreach (DataRow row in fullMembersTable.Rows)
                {
                    if (username == row["Username"].ToString())
                    {
                        email = row["Email"].ToString();
                        teamname = row["Teamname1"].ToString();
                        seatcode = row["Seatcode"].ToString();
                        break;
                    }
                }

                // Digital Workplace
                DigitalPictureBox.BackColor = seatcode.Contains("DW") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // RSS
                RSSPictureBox.BackColor = seatcode.Contains("RS") ? Color.FromArgb(80, 201, 196, 169) : Color.Transparent;
                // Global Ops
                GlobalOpsPictureBox.BackColor = seatcode.Contains("GO") ? Color.FromArgb(80, 239, 211, 210) : Color.Transparent;
                // FI Application Hosting
                FIPictureBox.BackColor = seatcode.Contains("FI") ? Color.FromArgb(80, 211, 223, 238) : Color.Transparent;
                // BI + Infor
                BIPictureBox.BackColor = seatcode.Contains("BI") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
                // Foundation Services + App Ops
                FoundationPictureBox.BackColor = seatcode.Contains("FS") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
                // Process Admin HR
                HRPictureBox.BackColor = seatcode.Contains("HR") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;
                // IDA + NTS + MB
                IDAPictureBox.BackColor = seatcode.Contains("ID") ? Color.FromArgb(80, 223, 223, 223) : Color.Transparent;
                // Centralized Service Desk
                CSDPictureBox.BackColor = seatcode.Contains("SD") ? Color.FromArgb(80, 253, 224, 201) : Color.Transparent;
                // Temporary Place
                TempPictureBox.BackColor = seatcode.Contains("TP") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 1
                Mg1PictureBox.BackColor = seatcode.Contains("MG1") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 2
                Mg2PictureBox.BackColor = seatcode.Contains("MG2") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 3
                Mg3PictureBox.BackColor = seatcode.Contains("MG3") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Reception
                ReceptionPictureBox.BackColor = seatcode.Contains("RE") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;


                int indexTeam;
                int indexUser;

                indexTeam = TeamsComboBox.Items.IndexOf(teamname);

                //if (indexTeam != -2 || indexTeam == -1)
                //{
                //    TeamsComboBox.SelectedIndexChanged -= new EventHandler(TeamsComboBox_SelectedIndexChanged);
                //}

                TeamsComboBox.SelectedIndex = indexTeam;
                TeamsComboBox.SelectedValue = teamname;

                DataTable membersTable = _dataSetSource.GetTeamMembers(teamname);
                //List<Member> membersList = membersTable.AsEnumerable().Select(dataRow => new Member(dataRow.Field<string>("Username"), dataRow.Field<string>("FullName"), dataRow.Field<string>("Sip"), dataRow.Field<string>("Email"), dataRow.Field<string>("Phone"), dataRow.Field<string>("Mobile"), dataRow.Field<string>("Manager"), dataRow.Field<string>("Department"), dataRow.Field<string>("Title"), dataRow.Field<string>("TeamName1"), dataRow.Field<string>("TeamName2"), dataRow.Field<string>("TeamName3"), dataRow.Field<string>("SeatCode"), dataRow.Field<string>("IsTeamLead"), dataRow.Field<string>("IsManager"), dataRow.Field<string>("IsOwner"))).Where(x => x.TeamName1 == teamname || x.TeamName2 == teamname || x.TeamName3 == teamname).ToList();

                List<Member> membersList = membersTable
                    .AsEnumerable()
                    .Select(dataRow => new Member(
                    dataRow.Field<string>("Username"),
                    dataRow.Field<string>("FullName"),
                    dataRow.Field<string>("Sip"),
                    dataRow.Field<string>("Email"),
                    dataRow.Field<string>("Phone"),
                    dataRow.Field<string>("Mobile"),
                    dataRow.Field<string>("Manager"),
                    dataRow.Field<string>("Department"),
                    dataRow.Field<string>("Title"),
                    dataRow.Field<string>("TeamName1"),
                    dataRow.Field<string>("TeamName2"),
                    dataRow.Field<string>("TeamName3"),
                    dataRow.Field<string>("SeatCode"),
                    dataRow.Field<string>("IsTeamLead"),
                    dataRow.Field<string>("IsManager"),
                    dataRow.Field<string>("IsOwner"),
                    dataRow.Field<string>("IsDirector")))
                    .Where(x => x.TeamName1 == teamname || x.TeamName2 == teamname || x.TeamName3 == teamname)
                    .OrderBy(x => x.Username)
                    .OrderBy(x => x.IsTeamLead == false)
                    .OrderBy(x => x.IsOwner == false)
                    .OrderBy(x => x.IsManager == false)
                    .OrderBy(x => x.IsDirector == false)
                    .ToList();

                _membersBindingSource.DataSource = membersList;
                MembersListBox.DataSource = _membersBindingSource;
                MembersListBox.DisplayMember = "Username ASC";
                MembersListBox.ValueMember = "Username";

                indexUser = MembersListBox.Items.IndexOf("username");

                MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);

                if (MembersListBox.Items.Count > 0)
                {
                    MembersListBox.SelectedIndex = indexUser;
                    MembersListBox.SelectedValue = username;
                }

                //ResetAreas();
                //SearchComboBox.Items.IndexOf(username);


                // initialize presence on buttons
                foreach (DataRow row in fullMembersTable.Rows)
                {
                    sip = row["Sip"].ToString();
                    _contact = _client.ContactManager.GetContactByUri(sip);
                    // initialize buttons
                    seatcode = row["Seatcode"].ToString();
                    if (!string.IsNullOrEmpty(seatcode))
                    {
                        foundButton = Controls.Find(seatcode, true);

                        if (foundButton != null && foundButton.Length > 0)
                        {
                            button = foundButton[0] as RoundButton;
                        }
                    }
                    activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();
                    switch (activity)
                    {
                        case "Available":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.LimeGreen;
                                }));
                            }
                            break;
                        case "Busy":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Red;
                                }));
                            }
                            break;
                        case "Do not disturb":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Red;
                                }));
                            }
                            break;
                        case "In a meeting":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Red;
                                }));
                            }
                            break;
                        case "In a call":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Red;
                                }));
                            }
                            break;
                        case "Be right back":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Gold;
                                }));
                            }
                            break;
                        case "Off work":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Gold;
                                }));
                            }
                            break;
                        case "Away":
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.Gold;
                                }));
                            }
                            break;
                        default:
                            if (button != null)
                            {
                                Invoke(new Action(() =>
                                {
                                    button.BackColor = Color.LightGray;
                                }));
                            }
                            break;
                    }
                }

                // highlight currently selected button
                seatcode = SeatCodeTextBox.Text;

                Invoke(new Action(() =>
                {
                    if (!string.IsNullOrEmpty(seatcode))
                    {
                        foundButton = Controls.Find(seatcode, true);
                        button = foundButton[0] as RoundButton;
                        button.BackColor = Color.Aqua;
                    }
                }));

                if (TabPanel.SelectedTab.Text == "Grid")
                {
                    MembersListBox.SelectedIndexChanged -= new EventHandler(MembersListBox_SelectedIndexChanged);
                    // populate Activity field in MembersDataGridView
                    for (int i = 0; i < MembersListBox.Items.Count; i++)
                    {
                        sip = MembersDataGridView.Rows[i].Cells["Sip"].Value.ToString();
                        _contact = _client.ContactManager.GetContactByUri(sip);
                        activity = _contact.GetContactInformation(ContactInformationType.Activity).ToString();

                        switch (activity)
                        {
                            case "Available":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.LimeGreen;
                                }));
                                break;
                            case "Busy":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                }));
                                break;
                            case "Do not disturb":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                }));
                                break;
                            case "In a call":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                }));
                                break;
                            case "In a meeting":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Red;
                                }));
                                break;
                            case "Be right back":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                }));
                                break;
                            case "Off work":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                }));
                                break;
                            case "Away":
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.DarkGoldenrod;
                                }));
                                break;
                            default:
                                Invoke(new Action(() =>
                                {
                                    MembersDataGridView.Rows[i].Cells["Activity"].Style.ForeColor = Color.Black;
                                }));
                                break;
                        }
                        Invoke(new Action(() =>
                        {
                            MembersDataGridView.Rows[i].Cells["Activity"].Value = activity;
                        }));
                    }
                    MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);
                }
                TeamsComboBox.SelectedIndexChanged += new EventHandler(TeamsComboBox_SelectedIndexChanged);
                SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            }
            else
            {
                TeamsComboBox.SelectedIndexChanged += new EventHandler(TeamsComboBox_SelectedIndexChanged);
                TeamsComboBox.SelectedIndex = 0;

                FullNameTextBox.Clear();
                TitleTextBox.Clear();
                SeatCodeTextBox.Clear();
                EmailTextBox.Clear();
                DepartmentTextBox.Clear();
                TeamTextBox.Clear();
                ActivityTextBox.Clear();
                ManagerCheckBox.Checked = false;
                ServiceOwnerCheckBox.Checked = false;
                TeamLeaderCheckBox.Checked = false;

                ContactPhotoBox.Refresh();
                // load default profile pic
                using (Stream photoStream = myAssembly.GetManifestResourceStream("TeamNavigator.Resources.default_profile_pic.jpg") as Stream)
                {
                    if (photoStream != null)
                    {
                        Bitmap bm = new Bitmap(photoStream);
                        ContactPhotoBox.Image = bm;
                    }
                }

                Invoke(new Action(() =>
                {
                    ContactBackPhotoBox.BackColor = Color.LightGray;
                }));
                
            }
            
            //SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            //MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);

        }
        private void DisplayUserInfo(object sender)
        {
            ToolTip tp = new ToolTip();
            string teamstring = "";
            string username = "";
            string fullname = "";
            foreach (DataRow row in fullMembersTable.Rows)
            {
                if (row["SeatCode"].ToString() == ((RoundButton)sender).Name)
                {
                    teamstring = row["TeamName1"].ToString() + "\r\n";
                    if (row["TeamName2"].ToString().Length > 0)
                    {
                        teamstring += row["TeamName2"].ToString() + "\r\n";
                    }
                    if (row["TeamName3"].ToString().Length > 0)
                    {
                        teamstring += row["TeamName3"].ToString() + "\r\n";
                    }
                    username = row["Username"].ToString();
                    fullname = row["FullName"].ToString();
                }
            }
            if (!string.IsNullOrEmpty(username))
            {
                tp.SetToolTip((RoundButton)sender, $"{username}\r\n{fullname}\r\n{teamstring}{((RoundButton)sender).Name}");
            }
            else
            {
                tp.SetToolTip((RoundButton)sender, $"{((RoundButton)sender).Name}");
            }
            
        }
        private void GoToMember(object sender)
        {
           
            if (!vacantseats.Contains((RoundButton)sender))
            {
                MembersListBox.SelectedIndexChanged -= new EventHandler(MembersListBox_SelectedIndexChanged);
                SearchComboBox.SelectedIndexChanged -= new EventHandler(SearchComboBox_SelectedIndexChanged);
                string teamname = "";
                string username = "";
                foreach (DataRow row in fullMembersTable.Rows)
                {
                    if (row["SeatCode"].ToString() == ((RoundButton)sender).Name)
                    {
                        username = row["Username"].ToString();
                        teamname = row["Teamname1"].ToString();
                        seatcode = row["Seatcode"].ToString();
                        break;
                    }
                }

                // Digital Workplace
                DigitalPictureBox.BackColor = seatcode.Contains("DW") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // RSS
                RSSPictureBox.BackColor = seatcode.Contains("RS") ? Color.FromArgb(80, 201, 196, 169) : Color.Transparent;
                // Global Ops
                GlobalOpsPictureBox.BackColor = seatcode.Contains("GO") ? Color.FromArgb(80, 239, 211, 210) : Color.Transparent;
                // FI Application Hosting
                FIPictureBox.BackColor = seatcode.Contains("FI") ? Color.FromArgb(80, 211, 223, 238) : Color.Transparent;
                // BI + Infor
                BIPictureBox.BackColor = seatcode.Contains("BI") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
                // Foundation Services + App Ops
                FoundationPictureBox.BackColor = seatcode.Contains("FS") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
                // Process Admin HR
                HRPictureBox.BackColor = seatcode.Contains("HR") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;
                // IDA + NTS + MB
                IDAPictureBox.BackColor = seatcode.Contains("ID") ? Color.FromArgb(80, 223, 223, 223) : Color.Transparent;
                // Centralized Service Desk
                CSDPictureBox.BackColor = seatcode.Contains("SD") ? Color.FromArgb(80, 253, 224, 201) : Color.Transparent;
                // Temporary Place
                TempPictureBox.BackColor = seatcode.Contains("TP") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 1
                Mg1PictureBox.BackColor = seatcode.Contains("MG1") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 2
                Mg2PictureBox.BackColor = seatcode.Contains("MG2") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Management 3
                Mg3PictureBox.BackColor = seatcode.Contains("MG3") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
                // Reception
                ReceptionPictureBox.BackColor = seatcode.Contains("RE") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;

                int indexTeam = -2;
                int indexUser = -2;

                indexTeam = TeamsComboBox.Items.IndexOf(teamname);

                if (indexTeam != -2)
                {
                    TeamsComboBox.SelectedIndexChanged -= new EventHandler(TeamsComboBox_SelectedIndexChanged);
                }

                TeamsComboBox.SelectedIndex = indexTeam;
                TeamsComboBox.SelectedValue = teamname;

                DataTable membersTable = _dataSetSource.GetTeamMembers(teamname);
                List<Member> membersList = membersTable.AsEnumerable().Select(dataRow => new Member(
                    dataRow.Field<string>("Username"), 
                    dataRow.Field<string>("FullName"), 
                    dataRow.Field<string>("Sip"), 
                    dataRow.Field<string>("Email"), 
                    dataRow.Field<string>("Phone"), 
                    dataRow.Field<string>("Mobile"), 
                    dataRow.Field<string>("Manager"), 
                    dataRow.Field<string>("Department"), 
                    dataRow.Field<string>("Title"), 
                    dataRow.Field<string>("TeamName1"), 
                    dataRow.Field<string>("TeamName2"), 
                    dataRow.Field<string>("TeamName3"), 
                    dataRow.Field<string>("SeatCode"), 
                    dataRow.Field<string>("IsTeamLead"), 
                    dataRow.Field<string>("IsManager"),
                    dataRow.Field<string>("IsOwner"),
                    dataRow.Field<string>("IsDirector")))
                    .Where(x => x.TeamName1 == teamname || x.TeamName2 == teamname || x.TeamName3 == teamname)
                    .OrderBy(x => x.Username)
                    .OrderBy(x => x.IsTeamLead == false)
                    .OrderBy(x => x.IsOwner == false)
                    .OrderBy(x => x.IsManager == false)
                    .OrderBy(x => x.IsDirector == false)
                    .ToList();

                _membersBindingSource.DataSource = membersList;
                MembersListBox.DataSource = _membersBindingSource;
                MembersListBox.DisplayMember = "Username ASC";
                MembersListBox.ValueMember = "Username";

                indexUser = MembersListBox.Items.IndexOf("username");

                MembersListBox.SelectedIndexChanged += new EventHandler(MembersListBox_SelectedIndexChanged);

                if (MembersListBox.Items.Count > 0)
                {
                    MembersListBox.SelectedIndex = indexUser;
                    MembersListBox.SelectedValue = username;
                }

                //ResetAreas();

                TeamsComboBox.SelectedIndexChanged += new EventHandler(TeamsComboBox_SelectedIndexChanged);
                SearchComboBox.SelectedIndexChanged += new EventHandler(SearchComboBox_SelectedIndexChanged);
            }
            Focus();
        }
        private void MembersDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Clipboard.SetDataObject(MembersDataGridView.CurrentCell.Value.ToString(), false);
        }
        private void VacantButton_Click(object sender, EventArgs e)
        {

            string buttonText = VacantButton.Text;

            if (buttonText == "Show Vacant")
            {
                VacantButton.Text = "Hide Vacant";
                foreach (RoundButton button in vacantseats)
                {
                    button.Visible = true;
                }
            }

            if (buttonText == "Hide Vacant")
            {
                VacantButton.Text = "Show Vacant";
                foreach (RoundButton button in vacantseats)
                {
                    button.Visible = false;
                }
            }
            Focus();

        }
        private void TeamNavigatorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // unsubscribe at form close
            foreach (Contact contact in contactsList)
            {
                contact.ContactInformationChanged -= new EventHandler<ContactInformationChangedEventArgs>(PeerContact_ContactInformationChanged);
            }
        }
        private void ResetAreas()
        {
            // Digital Workplace
            DigitalPictureBox.BackColor = SeatCodeTextBox.Text.Contains("DW") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
            // RSS
            RSSPictureBox.BackColor = SeatCodeTextBox.Text.Contains("RS") ? Color.FromArgb(80, 201, 196, 169) : Color.Transparent;
            // Global Ops
            GlobalOpsPictureBox.BackColor = SeatCodeTextBox.Text.Contains("GO") ? Color.FromArgb(80, 239, 211, 210) : Color.Transparent;
            // FI Application Hosting
            FIPictureBox.BackColor = SeatCodeTextBox.Text.Contains("FI") ? Color.FromArgb(80, 211, 223, 238) : Color.Transparent;
            // BI + Infor
            BIPictureBox.BackColor = SeatCodeTextBox.Text.Contains("BI") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
            // Foundation Services + App Ops
            FoundationPictureBox.BackColor = SeatCodeTextBox.Text.Contains("FS") ? Color.FromArgb(80, 224, 234, 203) : Color.Transparent;
            // Process Admin HR
            HRPictureBox.BackColor = SeatCodeTextBox.Text.Contains("HR") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;
            // IDA + NTS + MB
            IDAPictureBox.BackColor = SeatCodeTextBox.Text.Contains("ID") ? Color.FromArgb(80, 223, 223, 223) : Color.Transparent;
            // Centralized Service Desk
            CSDPictureBox.BackColor = SeatCodeTextBox.Text.Contains("SD") ? Color.FromArgb(80, 253, 224, 201) : Color.Transparent;
            // Temporary Place
            TempPictureBox.BackColor = SeatCodeTextBox.Text.Contains("TP") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
            // Management 1
            Mg1PictureBox.BackColor = SeatCodeTextBox.Text.Contains("MG1") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
            // Management 2
            Mg2PictureBox.BackColor = SeatCodeTextBox.Text.Contains("MG2") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
            // Management 3
            Mg3PictureBox.BackColor = SeatCodeTextBox.Text.Contains("MG3") ? Color.FromArgb(80, 170, 198, 234) : Color.Transparent;
            // Reception
            ReceptionPictureBox.BackColor = SeatCodeTextBox.Text.Contains("RE") ? Color.FromArgb(80, 242, 205, 237) : Color.Transparent;
        }

        //
        // AREAS PAINT EVENTS
        //

        private void DigitalPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Digital Workplace
            Point[] digitalPoints =
            {
                new Point(18, 22), new Point(251,22),
                new Point(251,60), new Point(306,60),
                new Point(306, 114), new Point(90, 114),
                new Point(90, 212), new Point(18, 212)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(digitalPoints);
                DigitalPictureBox.Region = new Region(gp);
            }
        }
        private void RSSPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // RSS

            Point[] rssPoints =
            {
                new Point(18,212), new Point(90,212),
                new Point(90,177), new Point(147,177),
                new Point(147,341), new Point(18,341)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(rssPoints);
                RSSPictureBox.Region = new Region(gp);
            }
        }
        private void GlobalOpsPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Global Ops

            Point[] globalPoints =
            {
                new Point(18, 341), new Point(168,341),
                new Point(168,368), new Point(239,368),
                new Point(239,564), new Point(93,564),
                new Point(93,636), new Point(18,636)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(globalPoints);
                GlobalOpsPictureBox.Region = new Region(gp);
            }
        }
        private void FIPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // FI Application Hosting

            Point[] fiPoints =
            {
                new Point(93,564), new Point(406,564),
                new Point(406,478), new Point(492,478),
                new Point(492,636), new Point(232,636),
                new Point(232,584), new Point(173,584),
                new Point(173,636), new Point(93,636)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(fiPoints);
                FIPictureBox.Region = new Region(gp);
            }
        }
        private void BIPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // BI + Infor

            Point[] biPoints =
            {
                new Point(370,404), new Point(506,404),
                new Point(506,477), new Point(370,477)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(biPoints);
                BIPictureBox.Region = new Region(gp);
            }
        }
        private void FoundationPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Foundation Services + App Ops

            Point[] digitalPoints =
            {
                new Point(536,564), new Point(687,564),
                new Point(687,460), new Point(790,460),
                new Point(790,636), new Point(536,636)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(digitalPoints);
                FoundationPictureBox.Region = new Region(gp);
            }
        }
        private void HRPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Process Admin HR

            Point[] digitalPoints =
            {
                new Point(687,324), new Point(790,324),
                new Point(790,460), new Point(687,460)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(digitalPoints);
                HRPictureBox.Region = new Region(gp);
            }
        }
        private void IDAPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // IDA + NTS + MB

            Point[] digitalPoints =
            {
                new Point(708,22), new Point(745,22),
                new Point(745,82), new Point(790,82),
                new Point(790,269), new Point(708,269)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(digitalPoints);
                IDAPictureBox.Region = new Region(gp);
            }
        }
        private void CSDPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Centralized Service Desk

            Point[] digitalPoints =
            {
                new Point(306,22), new Point(708,22),
                new Point(708,98), new Point(470,98),
                new Point(470,174), new Point(337,174),
                new Point(337,98), new Point(306,98)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(digitalPoints);
                CSDPictureBox.Region = new Region(gp);
            }
        }
        private void TempPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Temporary Place

            Point[] tempPoints =
            {
                new Point(147,177), new Point(250,177),
                new Point(250,250), new Point(147,250)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(tempPoints);
                TempPictureBox.Region = new Region(gp);
            }
        }
        private void ReceptionPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] receptionPoints =
{
                new Point(402,174), new Point(497,174),
                new Point(497,257), new Point(402,257)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(receptionPoints);
                ReceptionPictureBox.Region = new Region(gp);
            }
        }
        private void Mg1PictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] mg1Points =
{
                new Point(306,202), new Point(402,202),
                new Point(402,257), new Point(306,257)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(mg1Points);
                Mg1PictureBox.Region = new Region(gp);
            }
        }
        private void Mg2PictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] mg2Points =
{
                new Point(168,250), new Point(216,250),
                new Point(216,290), new Point(168,290)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(mg2Points);
                Mg2PictureBox.Region = new Region(gp);
            }
        }
        private void Mg3PictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] mg3Points =
{
                new Point(168,328), new Point(216,328),
                new Point(216,368), new Point(168,368)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(mg3Points);
                Mg3PictureBox.Region = new Region(gp);
            }
        }

        //
        // Rooms
        //

        private void RoomCasualPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] casualPoints =
{
                new Point(119,114), new Point(157,114),
                new Point(157,158), new Point(119,158)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(casualPoints);
                RoomCasualPictureBox.Region = new Region(gp);
            }
        }
        private void RoomYellowPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] yellowPoints =
{
                new Point(157,114), new Point(213,114),
                new Point(213,158), new Point(157,158)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(yellowPoints);
                RoomYellowPictureBox.Region = new Region(gp);
            }
        }
        private void RoomAquariumPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] aquariumPoints =
{
                new Point(251,22), new Point(306,22),
                new Point(306,61), new Point(251,61)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(aquariumPoints);
                RoomAquariumPictureBox.Region = new Region(gp);
            }
        }
        private void RoomLargePictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] largePoints =
{
                new Point(590,324), new Point(687,324),
                new Point(687,401), new Point(590,401)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(largePoints);
                RoomLargePictureBox.Region = new Region(gp);
            }
        }
        private void RoomGreenPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] greenPoints =
{
                new Point(590,401), new Point(687,401),
                new Point(687,448), new Point(590,448)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(greenPoints);
                RoomGreenPictureBox.Region = new Region(gp);
            }
        }
        private void RoomBluePictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] bluePoints =
{
                new Point(590,448), new Point(687,448),
                new Point(687,493), new Point(590,493)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(bluePoints);
                RoomBluePictureBox.Region = new Region(gp);
            }
        }
        private void RoomOneItPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] onePoints =
{
                new Point(517,448), new Point(590,448),
                new Point(590,493), new Point(517,493)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(onePoints);
                RoomOneItPictureBox.Region = new Region(gp);
            }
        }
        private void RoomVioletPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] violetPoints =
{
                new Point(517,401), new Point(590,401),
                new Point(590,448), new Point(517,448)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(violetPoints);
                RoomVioletPictureBox.Region = new Region(gp);
            }
        }
        private void Room2facesPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] facesPoints =
{
                new Point(492,564), new Point(536,564),
                new Point(536,636), new Point(492,636)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(facesPoints);
                Room2facesPictureBox.Region = new Region(gp);
            }
        }
        private void RoomCappuccinoPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] cappuccinoPoints =
{
                new Point(314,404), new Point(370,404),
                new Point(370,452), new Point(314,452)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(cappuccinoPoints);
                RoomCappuccinoPictureBox.Region = new Region(gp);
            }
        }
        private void GreyPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] greyPoints =
{
                new Point(259,404), new Point(314,404),
                new Point(314,452), new Point(259,452)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(greyPoints);
                GreyPictureBox.Region = new Region(gp);
            }
        }
        private void OrangePictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] orangePoints =
{
                new Point(173,584), new Point(232,584),
                new Point(232,636), new Point(173,636)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(orangePoints);
                OrangePictureBox.Region = new Region(gp);
            }
        }
        private void RedPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] redPoints =
{
                new Point(168,290), new Point(216,290),
                new Point(216,328), new Point(168,328)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(redPoints);
                RedPictureBox.Region = new Region(gp);
            }
        }
        private void RoomCornerPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Point[] cornerPoints =
{
                new Point(745,22), new Point(790,22),
                new Point(790,83), new Point(745,83)
            };

            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddPolygon(cornerPoints);
                RoomCornerPictureBox.Region = new Region(gp);
            }
        }

        //
        // MOUSE HOVER EVENTS
        //

        private void DW1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD20_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD26_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD23_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD22_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD21_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD18_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD16_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD17_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD15_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD32_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD30_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD27_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD28_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD29_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD31_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD35_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD34_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD33_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void ID13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS16_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS15_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS17_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS18_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS19_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS20_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS21_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS23_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS22_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI25_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI24_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI20_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI21_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI22_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI23_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI19_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI16_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI17_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI18_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI15_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI26_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI27_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI29_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI30_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI31_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI28_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void BI5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FI1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO28_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO27_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO26_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO23_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO24_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO25_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO22_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO21_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO20_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO19_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO18_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO17_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO16_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO15_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO29_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO30_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void TP1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void TP2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void TP3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void TP4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void TP5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW21_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW22_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW23_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW24_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS12_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS11_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS5_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS6_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS7_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS15_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS14_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS10_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS8_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RS16_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW17_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW18_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW20_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void DW19_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD13_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD24_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD25_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void HR3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void FS9_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void SD19_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO2B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO3B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO6B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO7B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO9B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO10B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO15B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO16B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO18B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO19B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO21B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO23B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO24B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void GO25B_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void MG1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void MG2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void MG3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RE1_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RE2_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RE3_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }
        private void RE4_MouseHover(object sender, EventArgs e)
        {
            DisplayUserInfo(sender);
        }

        //
        //
        // BUTTON CLICK
        //
        //

        private void DW1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD26_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD25_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD24_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD20_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD19_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD23_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD22_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD21_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD18_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD17_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD16_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD15_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD28_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD27_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD30_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD29_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD31_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD32_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD35_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD34_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void SD33_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void ID13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void HR7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS16_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS15_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS17_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS18_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS19_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS20_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS21_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS22_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FS23_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI25_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI24_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI20_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI21_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI22_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI23_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI19_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI16_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI17_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI18_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI26_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI27_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI28_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI29_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI30_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI31_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void BI5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI15_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void FI1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO28_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO27_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO26_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO23_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO23B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO24_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO24B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO25_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO25B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO22_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO21_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO21B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO20_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO19_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO19B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO18_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO18B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO17_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO15_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO15B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO16_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO16B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO7B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO2B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO3B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO30_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO29_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO9B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO10B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void GO6B_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS7_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS6_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS16_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS15_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS14_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS13_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS8_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS9_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS10_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS12_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RS11_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW24_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW23_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW22_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW21_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void TP1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void TP2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void TP3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void TP4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void TP5_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW19_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW20_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW18_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void DW17_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void MG1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void MG2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void MG3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RE1_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RE2_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RE4_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
        private void RE3_Click(object sender, EventArgs e)
        {
            GoToMember(sender);
        }
    }
}
