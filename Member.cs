
namespace TeamNavigator
{
    class Member
    {
        public bool IsDirector { get; set; }
        public bool IsTeamLead { get; set; }
        public bool IsManager { get; set; }
        public bool IsOwner { get; set; }
        public string Activity { get; set; }
        public string Username { get; set; }
        public string FullName { get; set; }
        public string Sip { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Mobile { get; set; }
        public string Manager { get; set; }
        public string Department { get; set; }
        public string Title { get; set; }
        public string TeamName1 { get; set; }
        public string TeamName2 { get; set; }
        public string TeamName3 { get; set; }
        public string SeatCode { get; set; }

        

        //public Member(string email)
        //{
        //    Email = email;
        //    Username = null;
        //    TeamName = null;
        //    Activity = null;
        //    Title = null;
        //    SeatCode = null;
        //    IsManager = false;
        //    IsTeamLead = false;
        //}

        public Member(string username, string fullname, string sip, string email, string phone, string mobile, string manager, string department, string title, string teamname1, string teamname2, string teamname3, string seatcode, string isteamlead, string ismanager, string isowner, string isdirector)
        {
            IsDirector = (isdirector == "Y") ? true : false;
            IsTeamLead = (isteamlead == "Y") ? true : false;
            IsManager = (ismanager == "Y") ? true : false;
            IsOwner = (isowner == "Y") ? true : false;
            Activity = "";
            Username = username;
            FullName = fullname;
            Sip = sip;
            Email = email;
            Phone = (phone != null) ? "+" + phone : "";
            Mobile = (mobile != null) ? "+" + mobile : "";
            Manager = (manager != null) ? manager : "";
            Department = (department != null) ? department : "";
            Title = (title != null) ? title : "";
            TeamName1 = teamname1;
            TeamName2 = (teamname2 != null) ? teamname2 : "";
            TeamName3 = (teamname3 != null) ? teamname3 : "";
            SeatCode = seatcode;
        }
    }
}
