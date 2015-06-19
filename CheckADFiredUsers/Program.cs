using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Configuration;

namespace CheckADFiredUsers
{
    class User
    {
        public string DisplayName { get; set; }
        public string Login { get; set; }
        public string EmployeeID { get; set; }
        public bool Fired { get; set; }
        public bool VPN { get; set; }
    }
    class Program
    {
        public static Int64 ConvertADSLargeIntegerToInt64(object adsLargeInteger)
        {
            var highPart = (Int32)adsLargeInteger.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, adsLargeInteger, null);
            var lowPart = (Int32)adsLargeInteger.GetType().InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, adsLargeInteger, null);
            return highPart * ((Int64)UInt32.MaxValue + 1) + lowPart;
        }
        static void Main(string[] args) 
        {
            Console.WriteLine(DateTime.Parse("25.03.2014 00:00:00").ToOADate());
            return;
            try
            {
                string p,p1,u;
                Console.WriteLine("Enter Login");
                u = Console.ReadLine();

                Console.WriteLine("Enter OLD PWD");
                Console.ForegroundColor = ConsoleColor.Black;
                Console.BackgroundColor = ConsoleColor.Black;
                p=Console.ReadLine();

                Console.ResetColor();
                Console.WriteLine("Enter NEW PWD");
                Console.ForegroundColor = ConsoleColor.Black;
                Console.BackgroundColor = ConsoleColor.Black;
                p1 = Console.ReadLine();
                Console.ResetColor();

                DirectoryEntry de1;
                
                using (de1 = new DirectoryEntry(ConfigurationManager.AppSettings["de1LDAPRoot"], u, p))
                {

                    DirectorySearcher deSearch = new DirectorySearcher();
                    deSearch.CacheResults = false;
                    deSearch.SearchRoot = de1;
                    deSearch.Filter = "(&(objectClass=user)(samAccountName="+u+"))";
                    SearchResult result = null; ;
                    try
                    {
                        result = deSearch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry deUser = new DirectoryEntry(result.Path);
                            Console.WriteLine(deUser.Properties["PasswordExpired"].Value);
                            /*               deUser.InvokeSet(
                            "AccountExpirationDate",
                            new object[] { DateTime.Now.AddDays(1) });
                            */
                            // Commit the changes.

                            // deUser.CommitChanges();

                        }
                        else { Console.WriteLine("Result is NULL"); }
                    }
                    catch (Exception e1)
                    {
                        Console.WriteLine("Enter OLD PWD");
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.BackgroundColor = ConsoleColor.Black;
                        p = Console.ReadLine();
                        Console.ResetColor();

                        de1 = new DirectoryEntry(ConfigurationManager.AppSettings["de1LDAPRoot"], ConfigurationManager.AppSettings["e1de1Login"], p);
                        deSearch = new DirectorySearcher();
                        deSearch.CacheResults = false;
                        deSearch.SearchRoot = de1;
                        deSearch.Filter = "(&(objectClass=user)(samAccountName=" + u + "))";
                        result = deSearch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry deUser = new DirectoryEntry(result.Path, ConfigurationManager.AppSettings["e1de1Login"], p);
                            // set up machine-level context
                            Console.WriteLine(deUser.Properties["PasswordExpired"].Value);

                            int m_Val1 = (int)deUser.Properties["userAccountControl"].Value;
                            int m_Val2 = (int)0x10000;
                            if (Convert.ToBoolean(m_Val1 & m_Val2))
                            {
                                Console.WriteLine("Password expired");
                            }

                            Int64 pls;
                            int uac;
                            pls = ConvertADSLargeIntegerToInt64(deUser.Properties["pwdLastSet"].Value);
                            uac = (int)deUser.Properties["UserAccountControl"].Value;

                            if ((pls == 0) && ((uac & 0x00010000) == 0) ? true : false)
                            {
                                Console.WriteLine("Password expired");
                                deUser.Properties["pwdLastSet"].Value = -1;
                                int val;
                                const int ADS_UF_DONT_EXPIRE_PASSWD = 0x10000;
                                val = (int)deUser.Properties["userAccountControl"].Value;
                                deUser.Properties["userAccountControl"].Value = val |
                                  ADS_UF_DONT_EXPIRE_PASSWD;
                               
                                deUser.CommitChanges();
                            }

                            // dirEntry.Properties["PasswordExpired"].Value = 0;
                            // dirEntry.CommitChanges();
                        }
                        else { Console.WriteLine("Result 2 is NULL"); }
                    }
                    // deUser.Invoke("SetPassword", new object[] { p1 });
                    //  deUser.Properties["LockOutTime"].Value = 0;
                    //  deUser.CommitChanges();
                }
                return;
           
                
                List<User> Users1C = new List<User>();
                Console.WriteLine();
                Console.WriteLine("Read 1C users");
                StreamReader sr = new StreamReader("1cusers.csv", Encoding.GetEncoding(1251));
                int rnum = 0;
                string user1clist;
                while ((user1clist = sr.ReadLine()) != null)
                {
                    if (rnum > 0) {  
                    Console.Write(".");
                    User user = new User();
                    user.DisplayName = user1clist.Split(';')[0];
                    user.EmployeeID = user1clist.Split(';')[1];
                    if (user1clist.Split(';')[4] != "01.01.2001 0:00" && user1clist.Split(';')[3] == "")
                    {
                        Console.WriteLine(user1clist.Split(';')[3]);
                    }
                    if (user1clist.Split(';')[4] == "01.01.2001 0:00" && user1clist.Split(';')[3] != "")
                    {
                        Console.WriteLine(user1clist.Split(';')[3]);
                    }
                    if (user1clist.Split(';')[4] != "01.01.2001 0:00" || user1clist.Split(';')[3] != "")
                    {
                     //   Console.WriteLine(user1clist);
                        if (DateTime.Parse(user1clist.Split(';')[4]) < DateTime.Parse(user1clist.Split(';')[5]))
                        {
                            Console.WriteLine(user1clist);
                        }
                        else
                        {
                            user.Fired = true;
                        }
                    }
                    
                    Users1C.Add(user);
                    }
                    rnum++;
                }
                Console.WriteLine();
                Console.WriteLine(Users1C.Count);

                List<User> UsersAD = new List<User>();
                DirectoryEntry de;
                using (de = new DirectoryEntry(ConfigurationManager.AppSettings["deLDAPRoot"]))
                {
                    Console.WriteLine();
                    Console.WriteLine("Read AD users");
                    DirectorySearcher deSearch = new DirectorySearcher();
                    deSearch.CacheResults = false;
                    deSearch.SearchRoot = de;
                    deSearch.Filter = "(&(objectClass=user))";// (cn=" + username + ")

                    SearchResultCollection result = deSearch.FindAll();
                    Console.WriteLine(result.Count);
                    foreach (SearchResult deu in result)
                    {

                        Console.Write(".");
                        DirectoryEntry deUser = new DirectoryEntry(deu.Path);
                        User user = new User();
                        user.DisplayName = (deUser.Properties["DisplayName"].Value ?? "").ToString();
                        user.Login=(deUser.Properties["sAMAccountName"].Value ?? "").ToString();
                        user.EmployeeID = (deUser.Properties["EmployeeID"].Value ?? "").ToString();

                        foreach (string GroupPath in deUser.Properties["memberOf"])
                        {
                            if (GroupPath.Equals(ConfigurationManager.AppSettings["vpnLDAPRoot"]))
                            {
                                user.VPN = true;
                            }
                        }

                        if (!string.IsNullOrEmpty(user.EmployeeID))
                        {
                            UsersAD.Add(user);
                        }
                        else
                        {
                            Console.WriteLine(user.DisplayName);
                            Console.WriteLine(user.EmployeeID);
                        }
                    }
                    Console.WriteLine();
                    Console.WriteLine(UsersAD.Count);
                }


                List<User> UsersSP = new List<User>();
                Console.WriteLine();
                Console.WriteLine("Read SP users");
                StreamReader spr = new StreamReader("VPN-sp.csv", Encoding.GetEncoding(1251));
                int spnum = 0;
                string usersplist;
                while ((usersplist = spr.ReadLine()) != null)
                {
                    if (spnum > 0)
                    {
                        Console.Write(".");
                        User user = new User();
                       // user.DisplayName = usersplist.Split(';')[0];
                        string lg=usersplist.Split(';')[0];
                        user.Login = lg.Substring(lg.ToLower().IndexOf(ConfigurationManager.AppSettings["dcname"]) + 4);
                        UsersSP.Add(user);
                    }
                    spnum++;
                }
                Console.WriteLine();
                Console.WriteLine(UsersSP.Count);


                foreach (User auser in UsersAD)
                {
                    bool found = false;
                    bool dupl = false;
                    bool nFired = false;
                    bool spfound = false;

                    foreach (User cuser in Users1C)
                    {
                        if (cuser.EmployeeID.Equals(auser.EmployeeID))
                        {
                            if (found)
                            {
                                dupl = true;

                            }
                            else
                            {
                                found = true;
                                if (cuser.Fired)
                                {
                                    nFired = true;
                                }
                            }
                        }
                    }
                    if (auser.VPN)
                    {
                        
                        foreach (User spuser in UsersSP)
                        {
                            if (spuser.Login.ToLower().Equals(auser.Login.ToLower()))
                            {
                             
                                spfound = true;
                            }                            
                        }
                    }
                    if (!found)
                    {
                        Console.WriteLine("Not Found: {0}; {1}", auser.DisplayName, auser.EmployeeID);
                    }
                    else if (dupl)
                    {
                        Console.WriteLine("Duplicate Found: {0}; {1}", auser.DisplayName, auser.EmployeeID);
                    }
                    else if (nFired)
                    {
                        Console.WriteLine("Not Fired Found: {0}; {1}", auser.DisplayName, auser.EmployeeID);
                    }
                    if (auser.VPN && !spfound)
                    {
                        Console.WriteLine("VPN Access Not Found on SP: {0}; {1}", auser.DisplayName, auser.EmployeeID);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
        }
    }
}
