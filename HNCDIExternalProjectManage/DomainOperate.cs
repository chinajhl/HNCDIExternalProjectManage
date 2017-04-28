using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace HNCDIExternalProjectManage
{
    class DomainOperate
    {
        private string stringDomainName;
        private DataTable arrayGroup;

        public DataTable ArrayGroup
        {
            get { return arrayGroup; }
            set { arrayGroup = value; }
        }
        private DataTable arrayUserName;

        public DataTable ArrayUserName
        {
            get { return arrayUserName; }
            set { arrayUserName = value; }
        }
        private DataTable arrayLoginID;

        public DataTable ArrayLoginID
        {
            get { return arrayLoginID; }
            set { arrayLoginID = value; }
        }

        public DomainOperate(string DomainName)
        {
            //
            //TODO: 在此处添加构造函数逻辑
            //
            stringDomainName = DomainName;
            entry = new DirectoryEntry("LDAP://" + DomainName);
            mySearcher = new DirectorySearcher(entry);
        }

        public DomainOperate()
        {
            //throw new System.NotImplementedException();
        }

        private DirectoryEntry entry;
        private DirectorySearcher mySearcher;

        public void GetGroup()
        {
            arrayGroup = new DataTable();
            //DataColumn group = arrayGroup.Columns.Add();
            //group.ColumnName = "value";
            //group.DataType = typeof(string);

            arrayGroup.Columns.Add("Value", typeof(string));
            arrayGroup.Columns.Add("Text", typeof(string));
            mySearcher.Filter = ("(objectClass=group)");
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                DirectoryEntry de = resEnt.GetDirectoryEntry();
                arrayGroup.Rows.Add(de.Properties["Name"].Value.ToString(), de.Properties["Name"].Value.ToString());
            }
        }

        private DataTable arrayOU;

        public DataTable ArrayOU
        {
            get { return arrayOU; }
            set { arrayOU = value; }
        }

        public void GetOU()
        {
            mySearcher.Filter = ("(objectClass=organizationalUnit)");
            arrayOU = new DataTable();
            arrayOU.Columns.Add("Value", typeof(string));
            arrayOU.Columns.Add("Text", typeof(string));
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                DirectoryEntry de = resEnt.GetDirectoryEntry();
                switch (de.Properties["Name"].Value.ToString())
                {
                    case "台式机":
                    case "笔记本":
                    case "计算机":
                    case "Domain Controllers":
                    case "temp":
                    case "Microsoft Exchange Security Groups":
                    case "tempOU":
                    case "tempou":
                    case "禁用账户":
                    case "金兴造价事务所":
                    case "外聘人员":
                    case "test":
                    case "组成员":
                    case "1.组成员":
                    case "服务器":
                    case "停用账户":
                    case "湖南省交通设计院":
                        continue;
                    default:
                        arrayOU.Rows.Add(de.Properties["Name"].Value.ToString(), de.Properties["Name"].Value.ToString());
                        break;
                }
            }
        }

        public void GetUsersByGroup(string stringGroup)
        {
            entry = new DirectoryEntry("LDAP://" + "CN=" + stringGroup + ",DC=" + stringDomainName + ",DC=com");
            mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = ("(objectClass=user)");
            arrayLoginID = new DataTable();
            arrayLoginID.Columns.Add("Value", typeof(string));
            arrayLoginID.Columns.Add("Text", typeof(string));
            arrayUserName = new DataTable();
            arrayUserName.Columns.Add("Value", typeof(string));
            arrayUserName.Columns.Add("Text", typeof(string));
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                try
                {
                    DirectoryEntry de = resEnt.GetDirectoryEntry();
                    arrayLoginID.Rows.Add(de.Properties["userPrincipalName"].Value.ToString(), de.Properties["userPrincipalName"].Value.ToString());
                    arrayUserName.Rows.Add(de.Properties["DisplayName"].Value.ToString(), de.Properties["DisplayName"].Value.ToString());
                }
                catch (Exception e)
                {
                    string strError = e.Message;
                }
            }
        }

        public void GetUsersByOU(string stringOU)
        {
            //entry = new DirectoryEntry("LDAP://" + "OU=" + stringOU + ",DC=" + stringDomainName + ",DC=com");
            entry = new DirectoryEntry("LDAP://" + "OU=湖南省交通设计院" + ",DC=" + stringDomainName + ",DC=com");
            DirectoryEntry ou = entry.Children.Find("OU=" + stringOU);
            
            mySearcher = new DirectorySearcher(ou);
            mySearcher.Filter = ("(objectClass=user)");
            arrayLoginID = new DataTable();
            arrayLoginID.Columns.Add("Value", typeof(string));
            arrayLoginID.Columns.Add("Text", typeof(string));
            arrayUserName = new DataTable();
            arrayUserName.Columns.Add("Value", typeof(string));
            arrayUserName.Columns.Add("Text", typeof(string));
            arrayUser = new DataTable();
            arrayUser.Columns.Add("Text", typeof(string));
            arrayUser.Columns.Add("Value", typeof(string));

            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                try
                {
                    DirectoryEntry de = resEnt.GetDirectoryEntry();
                    arrayLoginID.Rows.Add(de.Properties["userPrincipalName"].Value.ToString(), de.Properties["userPrincipalName"].Value.ToString());
                    arrayUserName.Rows.Add(de.Properties["userPrincipalName"].Value.ToString(), de.Properties["DisplayName"].Value.ToString());
                    arrayUser.Rows.Add(de.Properties["DisplayName"].Value.ToString(), de.Properties["userPrincipalName"].Value.ToString());
                }
                catch (Exception e)
                {
                    string strError = e.Message;
                }
            }
        }

        private DataTable arrayUser;

        public DataTable ArrayUser
        {
            get { return arrayUser; }
            set { arrayUser = value; }
        }

        public string GetUserNameByLoginID(string LoginID)
        {
            entry = new DirectoryEntry("LDAP://" + "DC=" + stringDomainName + ",DC=com");
            mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = "(&(objectClass=user)(userPrincipalName=" + LoginID + "))";
            SearchResult resEnt = mySearcher.FindOne();
            try
            {
                DirectoryEntry de = resEnt.GetDirectoryEntry();
                stringPassword = ""; // de.Properties["Password"].Value.ToString();
                return de.Properties["DisplayName"].Value.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        public List<string> GetLoginIDByUserName(string UserName)
        {
            entry = new DirectoryEntry("LDAP://" + "DC=" + stringDomainName + ",DC=com");
            mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = "(&(objectClass=user)(DisplayName=" + UserName + "))";
            List<string> logins = new List<string>();
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                DirectoryEntry de = resEnt.GetDirectoryEntry();
                stringPassword = ""; // de.Properties["Password"].Value.ToString();
                logins.Add(de.Properties["userPrincipalName"].Value.ToString());
            }
            return logins;
        }

        public string GetOuByLoginID(string LoginID)
        {
            entry = new DirectoryEntry("LDAP://" + "DC=" + stringDomainName + ",DC=com");
            mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = "(&(objectClass=user)(userPrincipalName=" + LoginID + "))";
            SearchResult resEnt = mySearcher.FindOne();
            try
            {
                DirectoryEntry de = resEnt.GetDirectoryEntry();
                return de.Parent.Properties["Name"].Value.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        private string stringPassword;

        public string StringPassword
        {
            get { return stringPassword; }
            set { stringPassword = value; }
        }

        private string stringMailBody;

        public string StringMailBody
        {
            get { return stringMailBody; }
            set { stringMailBody = value; }
        }

        public void SendMail()
        {
            //CDO.Message msg = new CDO.Message();

            //msg.From = stringLoginID;
            //msg.To = stringMailTo;
            //msg.Subject = "勘察设计资料互提卡";
            //msg.TextBody = stringMailBody;

            //CDO.IConfiguration iConfig = msg.Configuration;
            //ADODB.Fields fields = iConfig.Fields;

            //fields["http://schemas.microsoft.com/cdo/configuration/sendusing"].Value = 2;
            //fields["http://schemas.microsoft.com/cdo/configuration/sendemailaddress"].Value = stringLoginID;
            //fields["http://schemas.microsoft.com/cdo/configuration/sendpassword"].Value = stringPassword;
            //fields["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"].Value = 1;
            //fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"].Value = "131.100.200.228";

            //fields.Update();

            //try
            //{
            //    msg.Send();
            //    msg = null;
            //}
            //catch (Exception e)
            //{
            //    throw e;
            //}
            try
            {
                SmtpClient sMail = new SmtpClient("131.100.200.228");
                sMail.DeliveryMethod = SmtpDeliveryMethod.Network;
                sMail.Send(stringLoginID, stringMailTo, stringMailTitle, stringMailBody);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private string stringLoginID;

        public string StringLoginID
        {
            get { return stringLoginID; }
            set { stringLoginID = value; }
        }

        private string stringMailTo;

        public string StringMailTo
        {
            get { return stringMailTo; }
            set { stringMailTo = value; }
        }

        private string stringMailTitle;

        public string StringMailTitle
        {
            get { return stringMailTitle; }
            set { stringMailTitle = value; }
        }
    }
}
