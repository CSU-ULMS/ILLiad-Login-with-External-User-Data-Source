using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.DirectoryServices;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;

namespace Illiad
{
	public class Authentication
	{
		private string _PatronRecord;
		private string _FirstName;
		private string _LastName;
		private string _Barcode;
		private string _Username;
		private string _PatronType;
		private string _Department;
		private string _Email;
		private string _Phone;
    private string _StreetAddress;
    private string _Address2;
		private string _City;
		private string _State;
		private string _Zipcode;
		private string _MoneyOwed;
		private string _Status;
		private string _DepartmentAbbreviation;
    private string _ManualBlock;
		private DateTime _ExpirationDate;
		private DateTime _CatalogUpdatedDate;
		private DateTime _IlliadUpdatedDate;
				
		public bool CheckStandingLDAP( string strDomain, string strUsername, string strPassword )
		{
			/*
			 * Purpose:	To authenticate user against an LDAP source, like AD
			 * Accepts: strDomain = domain name
			 *			strUsername = user supplied username
			 *			strPassword = user supplied password
			 * Returns:	true if can bind to ldap source
			 *			throws error if cannot
			 */

			string strDomainUser = strDomain + @"\" + strUsername;
			DirectoryEntry objEntry = new DirectoryEntry("LDAP://" + strDomain, strDomainUser, strPassword);
					
			try
			{	
				//Bind to the native AdsObject to force authentication.			
				object objAuthenticate = objEntry.NativeObject;
		
				DirectorySearcher objSearch = new DirectorySearcher(objEntry);
		
				objSearch.Filter = "(SAMAccountName=" + strUsername + ")";
				objSearch.PropertiesToLoad.Add("cn");
				SearchResult objResult = objSearch.FindOne();
		
				if( objResult == null )
				{
					return false;
				}	
			}
			catch (Exception ex)
			{
				throw new Exception(System.Configuration.ConfigurationSettings.AppSettings["errorBadUsernamePassword"]);
			}
		
			return true;
		}

		public bool CheckStandingCatalog( string strKey, bool bolCampusUser )
		{
			/*
			 * Purpose:	Checks user's standing with catalog for ILL eligibility
			 * Accepts:	strKey = key to API (barcode, username or ssn, etc.)
			 *			bolCampusUser = true if CSUSM user, false if affiliate
			 * Returns:	true if success
			 *			throws error if failed
			 */

			string strError = "";

			// get data from catalog
			PullEntireRecord( strKey, bolCampusUser );

			// check good standing criteria
			if ( CheckPatronType( _PatronType ) != true )
			{
				strError += System.Configuration.ConfigurationSettings.AppSettings["errorNotAllowed"];
			}
			else
			{
				if ( CheckMoneyOwed( _MoneyOwed ) != true )
				{
					strError += System.Configuration.ConfigurationSettings.AppSettings["errorOwesMoney"];
				}
				if ( CheckForExpiration( _ExpirationDate ) != true )
				{
					strError += System.Configuration.ConfigurationSettings.AppSettings["errorExpired"];
				}
                if (CheckForManualBlock( _ManualBlock) != true)
                {
                    strError += System.Configuration.ConfigurationSettings.AppSettings["errorHoldOnRecord"];
                }
            }

			// if all is good, otherwise log error, throw exception

			if ( strError == "" )
			{
				return true;
			}
			else
			{
				throw new Exception ( strError );
			}
		}

		public bool CheckStandingIlliad (string strUsername )
		{
			/*
			 * Purpose:	Checks if user is already in Illiad
			 * Accepts:	strUsername = username that should correspond with illiad username
			 * Returns:	true or false
			 */
	
			string strConnString;				// database connection string
			string strSQL;						// sql statement	

			// set connection and SQL query strings
			strConnString = System.Configuration.ConfigurationSettings.AppSettings["LocalDatabase"];
			strSQL = "SELECT lastchangeddate FROM users WHERE username = '" + strUsername.ToLower() + "'";
			
			// connect to database
			OleDbConnection objConnection = new OleDbConnection(strConnString);
			objConnection.Open();
			OleDbCommand objCommand = new OleDbCommand(strSQL, objConnection);
			OleDbDataReader objReader = objCommand.ExecuteReader(CommandBehavior.CloseConnection);
			

			if (objReader.HasRows)
			{
				// get barcode
				while ( objReader.Read() ) 
				{
					_IlliadUpdatedDate = objReader.GetDateTime(0);
				}
				
				// close reader
				objReader.Close();

				return true;
			}
			else
			{
				// close reader
				objReader.Close();

				return false;
			}
		}
		public bool CheckAffiliateStandingIlliad (string strUsername )
		{
			/*
			 * Purpose:	Checks if user is already in Illiad
			 * Accepts:	strUsername = username that should correspond with illiad username
			 * Returns:	true or false
			 */
	
			string strConnString;				// database connection string
			string strSQL;						// sql statement	

			// set connection and SQL query strings
			strConnString = System.Configuration.ConfigurationSettings.AppSettings["LocalDatabase"];
			strSQL = "SELECT lastchangeddate, username FROM users WHERE ssn = '" + strUsername.ToLower() + "'";
			
			// connect to database
			OleDbConnection objConnection = new OleDbConnection(strConnString);
			objConnection.Open();
			OleDbCommand objCommand = new OleDbCommand(strSQL, objConnection);
			OleDbDataReader objReader = objCommand.ExecuteReader(CommandBehavior.CloseConnection);
			

			if (objReader.HasRows)
			{
				// get barcode
				while ( objReader.Read() ) 
				{
					_IlliadUpdatedDate = objReader.GetDateTime(0);
					_Username = objReader.GetString(1);
				}
				
				// close reader
				objReader.Close();

				return true;
			}
			else
			{
				// close reader
				objReader.Close();

				return false;
			}
		}
		public void PullEntireRecord ( string strKey, bool bolCampusUser ) 
		{
			/*
			 * Purpose:	Pulls down the entire patron record from catalog api, 
			 *			breaks it into parts, assigns those	to object properties
			 * Accepts:	strKey = patron API key (barcode, username, etc.)
			 *			bolCampusUser = true if CSUSM user, false if affiliate
			 * Returns:	Nothing or error message if a problem occured
			 */
			
			string strResponse = "";		// HTML response of API dump

			string strFirstName = "";		// first name of user
			string strLastName = "";		// last name of user

			string strBarcode = "";			// patron barcode
			string strUsername = "";		// username in pac

			string strPatronType = "";		// patron type of user
			string strDept = "";			// user's department or major

			string strEmail = "";			// email address of user
			string strPhone = "";			// phone number of user
			string strAddress = "";     // complete address of user
      string strStreetAddress = ""; // street address of user
      string strAddress2 = "";	// street address of user
			string strCity = "";			// city of user
			string strState = "";			// state of user
			string strZipcode = "";			// zipcode of user

			string strFines = "";			// fines amount
			string strUpdated = "";			// string version of date patron record in catalog updated
			string strExpDate = "";			// string version of date patron record expired in catalog
            string strMblock = "";           // any manual blocks or holds in pac record
			string strError = "";			// error message to user

			// normalize key if barcode
			strKey = strKey.Replace(" ", "");
			strKey = strKey.Replace("-", "");

			// get patron API dump from catalog
			strResponse = SendRequest(strKey);
            XmlDocument doc = new XmlDocument();
            doc.Load(strResponse);

            // Get xml nodes for each element in user data
            XmlNodeList itemNodes = doc.SelectNodes("//entry");
            foreach (XmlNode itemNode in itemNodes)
            {
                XmlNode fname = itemNode.SelectSingleNode("fname");
                XmlNode lname = itemNode.SelectSingleNode("lname");
                XmlNode UNIV_ID = itemNode.SelectSingleNode("UNIV_ID");
                XmlNode otherid1 = itemNode.SelectSingleNode("OTHER_ID_1");
                XmlNode almabarcode = itemNode.SelectSingleNode("BARCODE");
                XmlNode username = itemNode.SelectSingleNode("username");
                XmlNode groupid = itemNode.SelectSingleNode("groupid");
                XmlNode email = itemNode.SelectSingleNode("email");
                XmlNode phone = itemNode.SelectSingleNode("phone");
                XmlNode streetaddr1 = itemNode.SelectSingleNode("streetaddr1");
                XmlNode streetaddr2 = itemNode.SelectSingleNode("streetaddr2");
                XmlNode city = itemNode.SelectSingleNode("city");
                XmlNode state = itemNode.SelectSingleNode("state");
                XmlNode zipcode = itemNode.SelectSingleNode("zipcode");
                XmlNode blockstatus = itemNode.SelectSingleNode("blockstatus");
                XmlNode deptid = itemNode.SelectSingleNode("deptid");
                XmlNode fines = itemNode.SelectSingleNode("fines");
                XmlNode expdate = itemNode.SelectSingleNode("expdate");
                XmlNode moddate = itemNode.SelectSingleNode("moddate");

			    if ( strResponse.IndexOf("Requested record not found") != -1 )
			    {
				    strError += System.Configuration.ConfigurationSettings.AppSettings["errorNoPacRecord"];
			    }
			    else	
			    {

				    // assign entire response to PatronRecord field
				    _PatronRecord = strResponse;
				
				    // parse out individual fields with regular expression object

				    // PATRON NAME

				    strLastName = lname.InnerXml;
				    strFirstName = fname.InnerXml;
                    if (strLastName == "" || strFirstName == "")
				    {
				    	strError += System.Configuration.ConfigurationSettings.AppSettings["errorBadName"];
				    }
				    // BARCODE

                    if (UNIV_ID != null)
                    {
                        strBarcode = UNIV_ID.InnerXml;
                    }
                    else if (almabarcode != null)
                    {
                        strBarcode = almabarcode.InnerXml;
                    }
                    else if (otherid1 != null)
                    {
                        strBarcode = otherid1.InnerXml;
                    }
                    else
                    {
                        strBarcode = username.InnerXml;
                    }

				    // USERNAME

				    strUsername = username.InnerXml;

				    // PATRON TYPE

				    strPatronType = groupid.InnerXml;
                    if (strPatronType == "")
				    {
				    	strError += System.Configuration.ConfigurationSettings.AppSettings["errorMissingPatronTypeField"];
				    }

                    //IC 2014-10-28: add next three lines for VP's who have no CSUSM email.
                    if (strPatronType == "8")
                    {
                        bolCampusUser = false;
                    }

				    // DEPARTMENT

				    strDept = deptid.InnerXml;
                    if (strDept == "")
                    {
                        strError += System.Configuration.ConfigurationSettings.AppSettings["errorMissingDepartment"];
                    }
				    // E-MAIL

				    strEmail = email.InnerXml;
                    if (strEmail != "")
			    	{
					    // if this is campus user, make sure they have csusm e-mail address
	    				// otherwise is affiliate user, so ignore

                        if (strEmail.IndexOf("csusm") == -1 && bolCampusUser == true && strPatronType != "6")
				    	{
				    		strError += System.Configuration.ConfigurationSettings.AppSettings["errorMissingCSUSMEmail"];
				    	}
			    	}
			    	else
			    	{
					    strError += System.Configuration.ConfigurationSettings.AppSettings["errorMissingEmail"];
			    	}

				    // PHONE 

				    strPhone = phone.InnerXml;
                    if (strPhone == ""){
                        // for some reason, phone is never null in our ILLiad database
					    strPhone = "No Record";
                    }

          // ADDRESS

          strStreetAddress = streetaddr1.InnerXml;
          strAddress2 = streetaddr2.InnerXml;
					strCity = city.InnerXml;
					strState = state.InnerXml;
					strZipcode = zipcode.InnerXml;

                    if ( (strStreetAddress == "") || (strCity == "") || (strState == "") || (strZipcode == "")) 
					{

						// this is a non-standard entry

						// if this is a student, then something has gone wrong with the Banner/PeopleSoft
						// data load or Circ Desk has manually made a change.

						if ( strPatronType == "0" || strPatronType == "6" || strPatronType == "11" || strPatronType == "20" )
						{
							strError += System.Configuration.ConfigurationSettings.AppSettings["errorBadAddress"];
						}
						else if ( strPatronType == "1" || strPatronType == "2" || strPatronType == "14" || strPatronType == "16" || strPatronType == "17" || strPatronType == "19" || strPatronType == "21" )
						{
							// this is a faculty or staff member, so let's just take the field
							// in whatever way it comes

                            strStreetAddress = streetaddr1.InnerXml;
							strCity = "San Marcos";
							strState = "CA";
							strZipcode = "92096";
						}
					}
				}

				// MONEY OWED

				strFines = fines.InnerXml;

                // MANUAL BLOCK ON RECORD

                strMblock = blockstatus.InnerXml;
                if (strMblock == "")
                {
                    strError += System.Configuration.ConfigurationSettings.AppSettings["errorMissingManualBlockField"];
                }

                // EXPIRATION DATE

				strExpDate = expdate.InnerXml;
				if (strExpDate == "")
				{
					strError += System.Configuration.ConfigurationSettings.AppSettings["errorBadExpirationDate"];
				}

				// LAST UPDATED

				strUpdated = moddate.InnerXml;
				if (strUpdated == "")
				{
					strError += System.Configuration.ConfigurationSettings.AppSettings["errorBadLastUpdatedDate"];
				}
			}
            

			// caught problem data

			if (strError != "" )
			{
				// throw exception
				throw new Exception ( strError );
			}
			else
			{
				_FirstName = strFirstName;
				_LastName = strLastName;
				_Barcode = strBarcode;
				_Username = strUsername;
				_PatronType = strPatronType;
				_Status = ConvertPatronTypeAbbreviation(_PatronType);
				_Department = strDept;
				_DepartmentAbbreviation = ConvertMajorCode(_Department);
				_Email = strEmail;
				_Phone = strPhone;
				_StreetAddress = strStreetAddress;
				_City = strCity;
				_State = strState;
				_Zipcode = strZipcode;
				_MoneyOwed = strFines;
                _ManualBlock = strMblock;

				if ( strExpDate != "" )
				{
					_ExpirationDate = Convert.ToDateTime(strExpDate.Replace("-","/"));
				}

				if ( strUpdated != "" )
				{
					_CatalogUpdatedDate = Convert.ToDateTime(strUpdated.Replace("-","/"));
				}
			}

		}

		public int UpdateUserIlliad( string strUsername )
		{
			/*
			 * Purpose:	Updates Illiad record with new information from PAC
			 * Accepts:	srUsername = username
			 * Returns:	number of rows affected, should be 1
			 */

			// connect to database
			string strConnString = System.Configuration.ConfigurationSettings.AppSettings["LocalDatabase"];
			OleDbConnection objConn = new OleDbConnection(strConnString);
			objConn.Open();

			// SQL statement
			string strSQL = "UPDATE users " +

							"SET FirstName = ?, " +
							"LastName = ?, " +
							"Phone = ?, " +
							"EmailAddress = ?, " +
              "Address = ?, " +
              "Address2 = ?, " +
							"City= ?, " +
							"State = ?, " +
							"Zip = ?, " +
							"Department = ?, " +
							"ExpirationDate = ?, " +
							"LastChangedDate = ?, " +
							"Status = ?, " +
							"UserRequestLimit = ?, " +
							"SSN = ? " +

							"FROM users " +
							"WHERE username = ? ";
					
			OleDbCommand objCommand = new OleDbCommand(strSQL, objConn);
			objCommand.Parameters.Add("FirstName", _FirstName);
			objCommand.Parameters.Add("LastName", _LastName);
			objCommand.Parameters.Add("Phone", _Phone);
			objCommand.Parameters.Add("EmailAddress", _Email);
      objCommand.Parameters.Add("Address", _StreetAddress);
      objCommand.Parameters.Add("Address2", Address2);
			objCommand.Parameters.Add("City", _City);
			objCommand.Parameters.Add("State", _State);
			objCommand.Parameters.Add("Zip", _Zipcode);
			objCommand.Parameters.Add("Department", _DepartmentAbbreviation);	
			objCommand.Parameters.Add("ExpirationDate", _ExpirationDate);
			objCommand.Parameters.Add("LastChangedDate", DateTime.Now);
			objCommand.Parameters.Add("Status", Status);
			objCommand.Parameters.Add("UserRequestLimit", UserRequestLimit);
			objCommand.Parameters.Add("SSN", Barcode);

			objCommand.Parameters.Add("Username", strUsername.ToLower());

			
			// execute SQL query
			int iUpdated = objCommand.ExecuteNonQuery();
	
			// clean-up
			objConn.Close();

			return iUpdated;

		}
		public int AddUserIlliad( string strUsername, string strDeliveryMethod)
		{
			/*
			 * Purpose:	Adds a new Illiad record with user information from PAC
			 * Accepts:	srUsername = username
			 *			strNotificationMethod = user selected notification method
			 *			strDeliveryMethod = user selected delivery method
			 * Returns:	number of rows affected, should be 1
			 */

			// connect to database
			string strConnString = System.Configuration.ConfigurationSettings.AppSettings["LocalDatabase"];
			OleDbConnection objConn = new OleDbConnection(strConnString);
			objConn.Open();

			// SQL statement
			string strSQL = "INSERT INTO users (FirstName, LastName, Username, SSN, " +
				"Phone, EmailAddress, Address, Address2, City, State, Zip, " +
				"Department, ExpirationDate, LastChangedDate, Status, UserRequestLimit, " +
				"ArticleBillingCategory, LoanBillingCategory, NVTGC, LoanDeliveryMethod, Password, Cleared, Web, " +
				"NotificationMethod, DeliveryMethod " +
				") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
					
			OleDbCommand objCommand = new OleDbCommand(strSQL, objConn);
			objCommand.Parameters.Add("FirstName", _FirstName);
			objCommand.Parameters.Add("LastName", _LastName);
			objCommand.Parameters.Add("Username", strUsername.ToLower());
			objCommand.Parameters.Add("SSN", Barcode);

			objCommand.Parameters.Add("Phone", _Phone);
			objCommand.Parameters.Add("EmailAddress", _Email);
      objCommand.Parameters.Add("Address", _StreetAddress);
      objCommand.Parameters.Add("Address2", _Address2);
			objCommand.Parameters.Add("City", _City);
			objCommand.Parameters.Add("State", _State);
			objCommand.Parameters.Add("Zip", _Zipcode);

			objCommand.Parameters.Add("Department", _DepartmentAbbreviation);	
			objCommand.Parameters.Add("ExpirationDate", _ExpirationDate);
			objCommand.Parameters.Add("LastChangedDate", DateTime.Now);
			objCommand.Parameters.Add("Status", Status);
			objCommand.Parameters.Add("UserRequestLimit", UserRequestLimit);

			objCommand.Parameters.Add("ArticleBillingCategory", "");	
			objCommand.Parameters.Add("LoanBillingCategory", "");	
			objCommand.Parameters.Add("NVTGC", "ILL");
			objCommand.Parameters.Add("LoanDeliveryMethod", "Hold for Pickup");
			objCommand.Parameters.Add("Password", "991231542182231222042291991493501221149119"); // 'yourpassword' encrypted
			objCommand.Parameters.Add("Cleared", "No");
			objCommand.Parameters.Add("Web", "Yes");

			objCommand.Parameters.Add("NotificationMethod", "Electronic");
			objCommand.Parameters.Add("DeliveryMethod", strDeliveryMethod);

			// execute SQL query
			int iUpdated = objCommand.ExecuteNonQuery();
	
			// clean-up
			objConn.Close();

            return iUpdated;
            // return objCommand.CommandText;

        }
        public int AddUserIlliadNotification(string strUsername)
        {
            /*
			 * Purpose:	Adds several Illiad records that indicate user's default notification choices
			 * Accepts:	srUsername = username
			 * Returns:	number of rows affected, should be length of ACTIVITY_TYPE
             * Programmer: Khuong Vu
			 */
            string[] ACTIVITY_TYPE = {"ClearedUser", "PasswordReset", "RequestCancelled", "RequestOther", "RequestOverdue", "RequestPickup", "RequestShipped", "RequestElectronicDelivery"};

            // connect to database
            string strConnString = System.Configuration.ConfigurationSettings.AppSettings["LocalDatabase"];
            OleDbConnection objConn = new OleDbConnection(strConnString);
            objConn.Open();

            // SQL statement
            string strSQL = "INSERT INTO UserNotifications (Username, ActivityType, NotificationType) VALUES (?, ?, ?)";

            // OleDbCommand objCommand = new OleDbCommand(strSQL, objConn);

            // execute SQL query
            int iUpdated = 0; int iCount = 0;

            foreach (string activityType in ACTIVITY_TYPE) {
                OleDbCommand objCommand = new OleDbCommand(strSQL, objConn);
                objCommand.Parameters.Add("Username", strUsername.ToLower());
                objCommand.Parameters.Add("ActivityType", activityType);
                objCommand.Parameters.Add("NotificationType", "Email");
                iUpdated = objCommand.ExecuteNonQuery();
                iCount = iCount + iUpdated;

            }

            // clean-up
            objConn.Close();

            return iCount;
            // return objCommand.CommandText;

		}
		public void WriteToLog (string strError, string strRecord, string strLog )
		{
			/*
			 * Purpose:	Writes errors to an RSS xml file
			 * Accepts:	strError = exception
			 *			strRecord = user record or identifier
			 *			strLog = which log to write to.
			 * Returns:	nothing, writes log to local file
			 */

			string strTitle = "";	// title for the error log

			// path to log file, rotated monthly with year-month on
			// end of file name

			string strPath = "";
			int iMonth = DateTime.Now.Month;
			int iYear = DateTime.Now.Year;
			string strYearMonth = iYear.ToString() + "-" + iMonth.ToString("00");

			if (strLog == "patronload" )
			{
				strPath = @"C:\inetpub\wwwroot\login\logs\patronload-" + strYearMonth + ".xml";
				strTitle = "ILLiad New User Registration Errors";
			}
			else if (strLog == "login")
			{
				strPath = @"C:\inetpub\wwwroot\login\logs\login-" + strYearMonth + ".xml";
				strTitle = "ILLiad Login Errors";
			}
		
			// if file does not exist create it

			if (!File.Exists(strPath)) 
			{
				using (StreamWriter objWriter = File.CreateText(strPath)) 
				{
					objWriter.WriteLine("<?xml version=\"1.0\"?>");
					objWriter.WriteLine("<rss version=\"2.0\">");
					objWriter.WriteLine("<channel>");
					objWriter.WriteLine("<title>" + strTitle + "</title>");
					objWriter.WriteLine("<description>Tracking all things bad for Illiad authentication.</description>");
					objWriter.WriteLine("<language>en-us</language>");
					objWriter.WriteLine("</channel>");
					objWriter.WriteLine("</rss>");
					
					objWriter.Close();
				}    
			}

			// load-up the document
			XmlDocument objXml = new XmlDocument();
			objXml.Load(strPath);

			// get channel and first item elements
			XmlNode channel = objXml.SelectSingleNode("/rss/channel");
			XmlNode firstItem = objXml.SelectSingleNode("/rss/channel/item");

			// item
			XmlElement item = objXml.CreateElement("item");
			channel.InsertBefore(item,firstItem);

			// title
			XmlElement title = objXml.CreateElement("title");
			title.InnerText = strRecord;
			item.AppendChild(title);

			// description
			XmlElement description = objXml.CreateElement("description");
			description.InnerText = strError;
			item.AppendChild(description);

			// pubDate
			XmlElement pubDate = objXml.CreateElement("pubDate");
			pubDate.InnerText = DateTime.Now.ToString();
			item.AppendChild(pubDate);

			// save out the changes
			XmlTextWriter objXmlWriter  = new XmlTextWriter(strPath,System.Text.Encoding.UTF8);
			objXmlWriter.Formatting = Formatting.Indented;
			objXml.Save(objXmlWriter);
			objXmlWriter.Close();

		}


		private string ConvertMajorCode( string strDeptNumber)
		{
			/*
			 * Purpose:	Converts ptype3 number to Illiad (Banner-like) major code
			 * Accepts:	String integer of patron type number
			 * Returns: Banner 2 to 4 letter code
			 * Notes:	Throughout the entire student/faculty feed process the major codes are 
			 *			consolidated and renamed.  When the feed comes in, some codes, such
			 *			as the various education majors, are already boiled down to a single
			 *			code 'EDUC'.  In other cases, we've consolidated codes at the catalog 
			 *			by assigning a single pcode3 for various majors, as with nursing.  In other 
			 *			cases we solidify various codes here under a single code, as with 'BUS'
			 *			for business.  Some Illiad codes differ from those in Banner, reflecting 
			 *			continued use of older major designations by the Library.
			 */

            int iDept;
            if (!int.TryParse(strDeptNumber, out iDept))
            {
                strDeptNumber = "99";
            }
            iDept = Convert.ToInt32(strDeptNumber);

			// convert all business disciplines to 'BUS'
			if ( iDept == 6 ||	// BAAC
				iDept == 7 ||	// BAFN
				iDept == 8 ||	// BAGB
				iDept == 9 ||	// BAHT
				iDept == 10 ||	// BAMG
				iDept == 11 ||	// BASS 
				iDept == 12 	// BAMK
				//iDept == 47 ||	// PBSU obsolete 2013
				//iDept == 48		// PMBA obsolete 2013
				)
			{
				iDept = 5;		// BUS
			}

			// convert all biology disciplines to 'BIO'
			if ( iDept == 2 ||	// BIOC
				iDept == 4		// BIOT
				)
			{
				iDept = 3;		// BIO
			}

			// convert all education disciplines to 'EDUC'
			//if ( iDept == 33 )	// LBST
			//{
			//	iDept = 3;		// EDUC //Error corrected Feb 2013
			//}

			// convert all nursing disciplines to 'NURS'
			if ( iDept == 41 || // NUVN
                iDept == 49 ) // NURP, NUPP
			{
				iDept = 40;		// NURS
			}

			// convert the Sociology disciplines to 'SOC'
			if ( iDept == 58 ) // SP
			{
				iDept = 59;		// SOC
			}

			// convert all Foreign Languages to 'FLA'
			if ( iDept == 23 || // FREN added Feb 2013
                iDept == 25 ) // GERM  added Feb 2013
			{
				iDept = 60;		// FLA inc SPAN too
			}

      switch ( iDept )
			{
				case 3 : return "BIO";
				case 5 : return "BUS";
				case 13 : return "CHEM";
        case 14 : return "COG"; //added Feb 2013
        case 15 : return "COM";		// Banner: COMM
				case 16 : return "CSC";		// Banner: CS
				case 17 : return "CRIM";
				case 19 : return "ECON";
				case 20 : return "EDUC";
        case 21 : return "BRST"; //added Feb 2013
        case 22 : return "ETHN"; //added Feb 2013
        case 24 : return "GEOG"; //added Feb 2013
        case 26 : return "GLOB"; //added Feb 2013
        case 27 : return "HIST";
				case 29 : return "HDEV";
				case 32 : return "KIN";		// Banner: KINE
        case 33 : return "LBST";  //LBST own category as of Feb 2013
        case 34 : return "LING"; //added Feb 2013
        case 35 : return "LTWR";
        case 36 : return "MASS"; //added Feb 2013
				case 37 : return "MATH";
        case 38 : return "NATV"; //added Feb 2013
				case 40 : return "NURS";
        case 43 : return "PHYS"; //added Feb 2013
				case 45 : return "PSCI";
				case 50 : return "PSY";		// Banner: PSYC
				case 56 : return "SSC";		// Banner: SS
				case 59 : return "SOC";
				case 60 : return "FLA";		// Banner: SPAN
				case 61 : return "SPE";		// Banner: SPEC
				case 65 : return "VPA";		// Banner: ARTS
				case 69 : return "WMST";
				case 70 : return "EDUC";    // For Joint Doctoral Program
				case 98 : return "UND";		// Changed from 63 Feb 2013	
				case 99 : return "NOM";		// ADDED Feb 2013 to catch gaps	


				// non-major department codes
				// these were established by resource sharing

				case 79 : return "ALCI";
				case 103 : return "EXST";
				case 118 : return "LIB";
				case 140 : return "UBP";

				default : return "NOM";
			}	
		}

		private string ConvertPatronTypeFull( string strTypeID )
		{
			/*
			 * Purpose:	Converts patron type number to readable string in patron reg form
			 * Accepts:	strTypeID = patron type identifier
			 * Returns:	User-readable text
			 */

			switch ( strTypeID )
			{
				case "0":	// Student -- Undergrad
					return "Undergraduate Student";

				case "1":	// Faculty
					return "Faculty";

				case "2":	// Staff
					return "Staff";

				case "6":	// Student -- Graduate
					return "Graduate Student";

				case "8":	// CSU Faculty
					return "Visiting Faculty";

				case "11":	// Student -- Distance Education
					return "Distance Education Undergraduate";

				case "12":	// Graduate CSU Long Beach Masters in Social Work (SE Add 3/20/06)
					return "CSULB Masters";

				case "14":	// Faculty -- Teaching Associate(SE Add 3/20/06)
					return "Teaching Associate";

				case "16":	// Library Staff
					return "Library Staff";

				case "17":	// Library Faculty
					return "Library Faculty";	
	
				case "19":	// Visiting Scholar
					return "Visiting Scholar";	
		
				case "20":	// Graduate -- Distance Education
					return "Distance Education Graduate";	
		
				case "21":	// Administrator
					return "Administrator";

        case "Faculty":  // Faculty
          return "Faculty";

        case "Administrator":  // Administrator 
          return "Administrator";

        case "Staff":  // Staff 
          return "Staff";

        case "Undergraduate":  // Undergraduate
          return "Undergraduate Student";

        case "GraduateStudent":  // Graduate Student
          return "Graduate Student";

        case "ExtendedEducation":  // Extended Education
          return "Extended Education";

        case "SpecialPrograms":  // Special Programs
          return "Special Programs";

        case "GraduateDistEd":  // Distance Education Graduate
          return "Distance Education Graduate";

        case "UndergraduateDistEd": // Distance Education Undergraduate
          return "Distance Education Undergraduate";

        case "LibInternalUse":  // Faculty-Staff Administrator
          return "Extended Education";

        case "VisitingScholar":  // Visiting Scholar
          return "Visiting Scholar";

        case "Retiree":  // Faculty-Staff Administrator
          return "Extended Education";

        case "TA/GA":  // Faculty-Staff Administrator
          return "Teaching Associate";
        default:
					return "";
			}
		}

		private string ConvertPatronTypeAbbreviation( string strTypeID )
		{
			/*
			 * Purpose:	Converts patron type number to Illiad abbreviations
			 * Accepts:	strTypeID = patron type identifier
			 */

			switch ( strTypeID )
			{
				case "0":	// Student -- Undergrad
					return "UG";

				case "1":	// Faculty
					return "FAC";

				case "2":	// Staff
					return "STAF";

				case "6":	// Student -- Graduate
					return "GRAD";

				case "8":	// CSU Faculty
					return "CSU";

				case "11":	// Student 
					return "DISED";

				case "12":	// Graduate CSU Long Beach Masters in Social Work (SE Add 3/20/06)
					return "GRAD";

				case "14":	// Faculty -- Teaching Associate (SE Add 3/20/06)
					return "FAC";

				case "16":	// Library Staff
					return "STAF";

				case "17":	// Library Faculty
					return "FAC";	
		
				case "19":	// Visiting Scholar
					return "VISSCH";	
		
				case "20":	// Dis Ed Grad Student
					return "DISGRAD";	
		
				case "21":	// Faculty_Staff Administrator [label change 8/2/07 SE]
					return "FSADM";

        case "Faculty":  // Faculty
          return "FAC";

        case "Administrator":  //  Administrator
          return "FSADM";

        case "Staff":  // Staff 
          return "STAF";

        case "Undergraduate":  // Undergraduate
          return "UG";

        case "GraduateStudent":  // GraduateStudent
          return "GRAD";

        case "GraduateDistEd":  // GraduateDistEd
          return "DISED";

        case "UndergraduateDistEd": // UndergraduateDistEd
          return "DISED";

        case "LibInternalUse":  // LibInternalUse
          return "STAF";

        case "VisitingScholar":  // VisitingScholar
          return "VISSCH";

        case "TA/GA":  // TA/GA
          return "GRAD";

        default:
					return "";
			}
		}

		private bool CheckPatronType( string strTypeID )
		{
			/*
			 * Purpose:	Checks to make sure the user's patron type matches accepted types
			 * Accepts:	strTypeID = patron type identifier
			 * Returns:	True for accepted patron types
			 */

			switch ( strTypeID )
			{
				case "0":	// Student -- Undergrad
					return true;

				case "1":	// Faculty
					return true;

				case "2":	// Staff
					return true;

				case "6":	// Student -- Graduate
					return true;

				case "8":	// CSU Faculty
					return true;

				case "11":	// Student Undergrad -- Distance Education
					return true;

				case "12":	// Graduate CSU Long Beach Masters in Social Work (SE Add 3/20/06)
					return true;

				case "14":	// Faculty -- Teaching Associate(SE Add 3/20/06)
					return true;

				case "16":	// Library Staff
					return true;

				case "17":	// Library Faculty
					return true;	
	
				case "19":	// Visiting Scholar
					return true;	
		
				case "20":	// Student Graduate -- Distance Education
					return true;

        case "21":  // Faculty-Staff Administrator
          return true;

        case "Faculty":  // Faculty
          return true;

        case "Administrator":  //  Administrator
          return true;

        case "Staff":  // Staff 
          return true;

        case "Undergraduate":  // Undergraduate
          return true;

        case "GraduateStudent":  // GraduateStudent
          return true;

        case "ExtendedEducation":  // ExtendedEducation
          return true;

        case "SpecialPrograms":  // SpecialPrograms
          return true;

        case "GraduateDistEd":  // GraduateDistEd
          return true;

        case "UndergraduateDistEd": // UndergraduateDistEd
          return true;

        case "LibInternalUse":  // LibInternalUse
          return true;

        case "VisitingScholar":  // VisitingScholar
          return true;

        case "Retiree":  // Retiree
          return true;

        case "TA/GA":  // TA/GA
          return true;

        default:
					return false;
			}
		}

        private bool CheckForManualBlock(string strBlock)
        {
            /*
             * Purpose: checks to see if library has place a hold on record
             * Accepts: strBlock = value from MBlock patron field
             * Returns: True if '-', which means not blocked
             */
            switch (strBlock)
            {
                case "-": // no block
                    return true;
                default:
                    return false;
            }
        }

        private bool CheckMoneyOwed ( string strAmount )
		{
			/*
			 * Purpose:	Checks to see if the current fines exceed the limit
			 * Accepts:	strAmount = string formatted dollar amount without dollar sign
			 *			is treated then simply as a floating point number
			 * Returns:	True if less than amount
			 */

			//double iAmount = Convert.ToDouble( strAmount );
			
			//if ( iAmount < 10 )
			//{
				return true;
			//}
			//else
			//{
				//return false;
			//}
		}
		private bool CheckForExpiration ( DateTime datExpiry )
		{
			/*
			 * Purpose:	Checks to see if the patron record has expired
			 * Accepts:	datExpiry = patron expiration date
			 * Returns:	True if user's record expiration date is greater than today
			 */

			// get current date
			DateTime datToday = DateTime.Now;
			
			// compare the two dates: -1 means the expiration date is older than today
			if ( datExpiry.CompareTo(datToday) == -1 )
			{
				return false;
			}
			else
			{
				return true;
			}
		}
		private string SendRequest ( string strKey ) 
		{
			/* 
			 * Purpose:	general function to send and recieve information via HTTP
			 * Accepts: strKey = the id (barcode, username) for the patron API
			 * Returns: HTML response
			 */
			
			string strResult = "";
			string PATRON_API = System.Configuration.ConfigurationSettings.AppSettings["PartonAPI"];
            strResult = PATRON_API += strKey;
			
			return strResult;
		}


		public string PatronRecord
		{
			get  
			{
				return _PatronRecord;
			}
		}
		public string FirstName
		{
			get  
			{
				return _FirstName;
			}
			set   
			{
				_FirstName = value;
			}
		}
		public string LastName
		{
			get  
			{
				return _LastName;
			}
			set   
			{
				_LastName = value;
			}
		}
		public string Barcode
		{
			get  
			{
				if ( _Barcode != "" && _Barcode != null )
				{
					return _Barcode;
				}
				else
				{
					return "12345678900000";
				}
			}
			set   
			{
				_Barcode = value;
			}
		}
		public string Username
		{
			get  
			{
				return _Username;
			}
			set   
			{
				_Username = value;
			}
		}
		public string PatronTypeID
		{
			get  
			{
				return _PatronType;
			}
			set   
			{
				_PatronType = value;
			}
		}
		public string PatronTypeFull
		{
			get  
			{
				return ConvertPatronTypeFull(_PatronType);
			}
		}
		public string Status
		{
			get  
			{
				return _Status;
			}
			set
			{
				_Status = value;
			}
		}
		public int UserRequestLimit
		{
			get  
			{
				if ( _Status == "FAC" || _Status == "GRAD" || _Status == "CSU")
				{
					return 30;
				}
				else
				{
					return 20;
				}
			}
		}
		public string DepartmentAbbreviation
		{
			get  
			{
				return _DepartmentAbbreviation;
			}
			set
			{
				_DepartmentAbbreviation = value;
			}
		}

		public string DepartmentID
		{
			get  
			{
				return _Department;
			}
			set   
			{
				_Department = value;
			}
		}

		public string Email
		{
			get  
			{
				return _Email;
			}
			set   
			{
				_Email = value;
			}
		}
		public string Phone
		{
			get  
			{
				return _Phone;
			}
			set   
			{
				_Phone = value;
			}
		}
    public string StreetAddress
    {
      get
      {
        return _StreetAddress;
      }
      set
      {
        _StreetAddress = value;
      }
    }
    public string Address2
		{
			get  
			{
				return _Address2;
			}
			set   
			{
        _Address2 = value;
			}
		}
		public string City
		{
			get  
			{
				return _City;
			}
			set   
			{
				_City = value;
			}
		}
		public string State	
		{
			get  
			{
				return _State;
			}
			set   
			{
				_State = value;
			}
		}
		public string Zipcode
		{
			get  
			{
				return _Zipcode;
			}
			set   
			{
				_Zipcode = value;
			}
		}

		public string MoneyOwed
		{
			get  
			{
				return _MoneyOwed;
			}
			set   
			{
				_MoneyOwed = value;
			}
		}

		public DateTime ExpirationDate
		{
			get  
			{
				return _ExpirationDate;
			}
			set   
			{
				_ExpirationDate = value;
			}
		}
		public DateTime CatalogUpdatedDate
		{
			get  
			{
				return _CatalogUpdatedDate;
			}
			set   
			{
				_CatalogUpdatedDate = value;
			}
		}
		public DateTime IlliadUpdatedDate
		{
			get  
			{
				return _IlliadUpdatedDate;
			}
			set   
			{
				_IlliadUpdatedDate = value;
			}
		}
	}
}
