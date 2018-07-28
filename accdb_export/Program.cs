using System;
using System.Linq;
using System.Data.OleDb;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using OfficeOpenXml.Style;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace accdb_export
{
    class Program
    {
        public static String source;

        public static String csv;
        public static String xlsx;

        //public static int product_number = -1;

        public static String like;
        public static List<int> product_numbers = new List<int>();
        public static String serial_string;

        static int LIKE_LOCATION = 0;
        static int LIKE_PRODUCT_NO = 1;
        static int LIKE_SERIAL_NO = 2;

        static int SEARCH_METHOD;

        public static String def_comp = "HL - Hamilton Library";

        static List<MicroRecord_Translated> items = new List<MicroRecord_Translated>();

        public static Settings SETTINGS { get; set; }
        public static JObject item_catalog { get; set; }

        //Copypasta'd directly from:
        //https://www.c-sharpcorner.com/article/read-microsoft-access-database-in-C-Sharp-6/
        static void Main(string[] args)
        {
            //TODO load settings
            if (File.Exists(".\\settings.json"))
            {
                SETTINGS = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(@"settings.json"));
                if (!File.Exists(SETTINGS.DEFAULT_PATH))
                {
                    SETTINGS.DEFAULT_PATH = "default.accdb";
                }

                if (File.Exists(SETTINGS.DEFAULT_ITEM_INDEX))
                {
                    item_catalog = JObject.Parse(JsonConvert.DeserializeObject(JsonConvert.SerializeObject(File.ReadAllText(SETTINGS.DEFAULT_ITEM_INDEX))).ToString());

                    foreach (var v in item_catalog)
                    {
                        Console.WriteLine(v.Key);
                    }
                } else
                {
                    //TODO create the file and tell user to edit it
                }
            }

            if (!args.Contains("/?"))
            {

                if (args.Contains("/s"))
                {
                    SETTINGS.DEFAULT_PATH = args[Array.IndexOf(args, "/s") + 1];
                }

                if (args.Contains("/loc"))
                {
                    like = args[Array.IndexOf(args, "/loc") + 1];

                    SEARCH_METHOD = LIKE_LOCATION;
                }
                else if (args.Contains("/pn"))
                {
                    String pn_args_string = args[Array.IndexOf(args, "/pn") + 1];

                    if (pn_args_string.Contains(','))
                    {
                        String[] pn_args = pn_args_string.Split(',');
                        foreach (String s in pn_args)
                        {
                            product_numbers.Add(Int32.Parse(s));
                        }
                    }
                    else
                    {
                        product_numbers.Add(Int32.Parse(pn_args_string));
                    }

                    SEARCH_METHOD = LIKE_PRODUCT_NO;
                }
                else if (args.Contains("/ser"))
                {
                    serial_string = args[Array.IndexOf(args, "/ser") + 1];

                    SEARCH_METHOD = LIKE_SERIAL_NO;
                }

                if (args.Contains("/xlsx"))
                {
                    xlsx = args[Array.IndexOf(args, "/xlsx") + 1];
                }
                else
                {
                    xlsx = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + ".xlsx";
                }
                if (args.Contains("/sl"))
                {
                    def_comp = "SL - Sinclair Library";
                }
                else if (args.Contains("/hl"))
                {
                    def_comp = "HL - Hamilton Library";
                }

                //import items

                //TODO UNCOMMENT
                int num = Database_Query();

                if (items.Count > 0)
                {
                    Export_Items();
                }

                Console.WriteLine("SOURCE=" + SETTINGS.DEFAULT_PATH);
                Console.WriteLine("LOCATION LIKE \"" + like + "\", PRODUCT_NUM LIKE \"" + product_numbers.ToString() + "\"");
                Console.WriteLine(num + " record(s) found.");

                if (args.Contains("/o") && items.Count > 0)
                {
                    Console.WriteLine("Opening file...");
                    System.Diagnostics.Process.Start(xlsx);
                }

            } else
            {
                String HELP_MSG = @"Micro .ACCDB to SnipeIT .XLSX Exporter v.1.0
Exports items from a .accdb file into a .xlsx format which is used to prep a csv for SnipeIT.

USAGE:
    accdb_export [ /? 
                    /loc [location] | 
                    /pn [1,2,n] | 
                    /ser [serial] |
                    /o ]

Options:
    /?              Display this help message
    /loc [location] Export items with ""LOCATION"" like [location]
    /pn [n0,n1,nx]     Export items with ""PRODUCT_NO"" like [n0,n1,nx]
    /ser [serial]   Export items with ""SERIAL_NUM"" or ""MONITOR"" like [serial]
    /o              Open exported file when done.
";

                Console.WriteLine(HELP_MSG);
            }

            //todo halt option
            //Console.WriteLine("Press any key to continue...");
            //Console.ReadKey();

        }

        static void Export_Items()
        {
            Console.WriteLine("Exporting...");

            ExcelPackage package = new ExcelPackage(new FileInfo(Path.GetFullPath(xlsx)));
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("LOCATION LIKE \"" + like + "\", PRODUCT_NUM LIKE \"" + product_numbers.ToString() + "\"");

            worksheet.Cells[1, 1].Value = "Full Name";
            worksheet.Cells[1, 2].Value = "Email";
            worksheet.Cells[1, 3].Value = "Username";
            worksheet.Cells[1, 4].Value = "Item Name";
            worksheet.Cells[1, 5].Value = "Category";
            worksheet.Cells[1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.Red);
            worksheet.Cells[1, 6].Value = "Model Name";
            worksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.Red);
            worksheet.Cells[1, 7].Value = "Manufacturer";
            worksheet.Cells[1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.Red);
            worksheet.Cells[1, 8].Value = "Model Number";
            worksheet.Cells[1, 9].Value = "Serial Number";
            worksheet.Cells[1, 10].Value = "Asset Tag";
            worksheet.Cells[1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 10].Style.Fill.BackgroundColor.SetColor(Color.Red);
            worksheet.Cells[1, 11].Value = "Location";
            worksheet.Cells[1, 12].Value = "Item Notes";
            worksheet.Cells[1, 12].Style.WrapText = true;
            worksheet.Cells[1, 13].Value = "Purchase Date";
            worksheet.Cells[1, 14].Value = "Purchase Cost";
            worksheet.Cells[1, 15].Value = "Company";
            worksheet.Cells[1, 16].Value = "Status";
            worksheet.Cells[1, 17].Value = "Warranty Months";
            worksheet.Cells[1, 18].Value = "Supplier";
            worksheet.Cells[1, 19].Value = "Order Number";
            worksheet.Cells[1, 20].Value = "Name";
            worksheet.Cells[1, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 20].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[1, 21].Value = "Dept/Room Location";
            worksheet.Cells[1, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 21].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[1, 22].Value = "Mac Address 1";
            worksheet.Cells[1, 23].Value = "Decal";
            worksheet.Cells[1, 24].Value = "Display Size";
            worksheet.Cells[1, 25].Value = "Touchscreen";
            worksheet.Cells[1, 26].Value = "RAM";

            int row = 2;
            foreach (MicroRecord_Translated i in items)
            {
                worksheet.Cells[row, 1].Value = i.Full_Name;
                worksheet.Cells[row, 2].Value = i.Email;
                worksheet.Cells[row, 3].Value = i.Username;
                worksheet.Cells[row, 4].Value = i.Item_Name;
                worksheet.Cells[row, 5].Value = i.Category;
                worksheet.Cells[row, 6].Value = i.Model_Name;
                worksheet.Cells[row, 7].Value = i.Manufacturer;
                worksheet.Cells[row, 8].Value = i.Model_Number;
                worksheet.Cells[row, 9].Value = i.Serial_Number;
                worksheet.Cells[row, 10].Value = i.Asset_Tag;
                worksheet.Cells[row, 11].Value = i.Location;
                worksheet.Cells[row, 12].Value = i.Notes;
                worksheet.Cells[row, 13].Value = i.Purchase_Date;
                worksheet.Cells[row, 14].Value = i.Purchase_Cost;
                worksheet.Cells[row, 15].Value = i.Company;
                worksheet.Cells[row, 16].Value = i.Status;
                worksheet.Cells[row, 17].Value = i.Warranty;
                worksheet.Cells[row, 18].Value = i.Supplier;
                worksheet.Cells[row, 19].Value = i.Order_Number;
                worksheet.Cells[row, 20].Value = i.Name;
                worksheet.Cells[row, 21].Value = i.Dept_Room_Location;
                worksheet.Cells[row, 22].Value = i.Mac_Address_1;
                worksheet.Cells[row, 23].Value = i.Decal;
                worksheet.Cells[row, 24].Value = i.Display_Size;
                worksheet.Cells[row, 25].Value = i.Touchscreen;
                worksheet.Cells[row, 26].Value = i.RAM;

                row++;
            }
            worksheet.Row(1).Style.Font.Bold = true;

            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();
            worksheet.Column(4).AutoFit();
            worksheet.Column(5).AutoFit();
            worksheet.Column(6).AutoFit();
            worksheet.Column(7).AutoFit();
            worksheet.Column(8).AutoFit();
            worksheet.Column(9).AutoFit();
            worksheet.Column(10).AutoFit();
            worksheet.Column(11).AutoFit();
            worksheet.Column(12).AutoFit();
            worksheet.Column(13).AutoFit();
            worksheet.Column(14).AutoFit();
            worksheet.Column(15).AutoFit();
            worksheet.Column(16).AutoFit();
            worksheet.Column(17).AutoFit();
            worksheet.Column(18).AutoFit();
            worksheet.Column(19).AutoFit();
            worksheet.Column(20).AutoFit();
            worksheet.Column(21).AutoFit();
            worksheet.Column(22).AutoFit();
            worksheet.Column(23).AutoFit();
            worksheet.Column(24).AutoFit();
            worksheet.Column(25).AutoFit();
            worksheet.Column(26).AutoFit();

            package.Save();
            Console.WriteLine(xlsx + " saved.");
        }
        

        //https://stackoverflow.com/questions/15128361/getting-data-from-ms-access-database-and-display-it-in-a-listbox
        static int Database_Query()
        {

            //https://stackoverflow.com/questions/16674024/ms-access-contains-query
            //string strSQL;
            //if (like.Equals("") && product_number == -1)
            //{
            //    strSQL = "SELECT * FROM Micro";
            //}
            //else if (!like.Equals("") && product_number == -1)
            //{
            //    strSQL = "SELECT * FROM Micro WHERE LOCATION Like '%" + like + "%'";
            //}
            //else if (like.Equals("") && !(product_number == -1))
            //{
            //    strSQL = "SELECT * FROM Micro WHERE PRODUCT_NO Like '%" + product_number + "%'";
            //}
            //else
            //{
            //    strSQL = "SELECT * FROM Micro WHERE LOCATION Like '%" + like + "%' AND PRODUCT_NO Like '%" + product_number + "%'";
            //}

            string strDSN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = '" + SETTINGS.DEFAULT_PATH + "'";

            int recordnum = 0;

            if (SEARCH_METHOD == LIKE_LOCATION)
            {
                String strSQL = "SELECT * FROM Micro WHERE LOCATION Like '%" + like + "%'";

                OleDbConnection myConn = new OleDbConnection(strDSN);
                myConn.Open();
                OleDbCommand cmd = new OleDbCommand(strSQL, myConn);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    MicroRecord_Translated item = new MicroRecord_Translated(reader);
                    item.Company = def_comp;
                    items.Add(item);
                    recordnum++;
                }
                myConn.Close();

            } else if (SEARCH_METHOD == LIKE_PRODUCT_NO)
            {
                foreach (int i in product_numbers)
                {
                    String strSQL = "SELECT * FROM Micro WHERE PRODUCT_NO Like '%" + i + "%'";

                    OleDbConnection myConn = new OleDbConnection(strDSN);
                    myConn.Open();
                    OleDbCommand cmd = new OleDbCommand(strSQL, myConn);
                    OleDbDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        MicroRecord_Translated item = new MicroRecord_Translated(reader);
                        item.Company = def_comp;
                        items.Add(item);
                        recordnum++;
                    }
                    myConn.Close();
                }
            } else if (SEARCH_METHOD == LIKE_SERIAL_NO)
            {
                String strSQL01 = "SELECT * FROM Micro WHERE SERIAL_NUM Like '%" + serial_string + "%'";

                OleDbConnection myConn01 = new OleDbConnection(strDSN);
                myConn01.Open();
                OleDbCommand cmd01 = new OleDbCommand(strSQL01, myConn01);
                OleDbDataReader reader01 = cmd01.ExecuteReader();

                while (reader01.Read())
                {
                    MicroRecord_Translated item = new MicroRecord_Translated(reader01);
                    item.Company = def_comp;
                    items.Add(item);
                    recordnum++;
                }
                myConn01.Close();

                String strSQL02 = "SELECT * FROM Micro WHERE MONITOR Like '%" + serial_string + "%'";

                OleDbConnection myConn02 = new OleDbConnection(strDSN);
                myConn02.Open();
                OleDbCommand cmd02 = new OleDbCommand(strSQL02, myConn02);
                OleDbDataReader reader02 = cmd02.ExecuteReader();

                while (reader02.Read())
                {
                    MicroRecord_Translated item = new MicroRecord_Translated(reader02);
                    item.Company = def_comp;
                    items.Add(item);
                    recordnum++;
                }
                myConn02.Close();
            }

            return recordnum;
        }

        class MicroRecord_Translated
        {
            public String Full_Name { get; set; }
            public String Email { get; set; }
            public String Username { get; set; }
            public String Item_Name { get; set; }
            public String Category { get; set; } //
            public String Model_Name { get; set; } //
            public String Manufacturer { get; set; } //
            public String Model_Number { get; set; }
            public String Serial_Number { get; set; }
            public String Asset_Tag { get; set; } //
            public String Location { get; set; }
            public String Notes { get; set; }
            public String Purchase_Date { get; set; }
            public String Purchase_Cost { get; set; }
            public String Company { get; set; }
            public String Status { get; set; }
            public String Warranty { get; set; }
            public String Supplier { get; set; }
            public String Order_Number { get; set; }
            public String Name { get; set; }
            public String Dept_Room_Location { get; set; }
            public String Mac_Address_1 { get; set; }
            public String Service_Tag { get; set; }
            public String Decal { get; set; }
            public String Display_Size { get; set; }
            public String Touchscreen { get; set; }
            public String RAM { get; set; }

            public String[][] Categories = new string[][] {
                new string[] { "APC", "Power - UPS" },
                new string[] { "UPS", "Power - UPS" },
                new string[] { "MONITOR", "Display - Monitor" },
                new string[] { "SCREEN", "Display - Monitor" },
                new string[] { "TV", "Display - TV" },
                new string[] { "LCD", "Display - Monitor" },
                new string[] { "P4015N", "Printer - Laser - Black and White" },
                new string[] { "P4015DN", "Printer - Laser - Black and White" },
                //new string[] { "PRT", "Printer - LaserJet - ???" },
                new string[] { "BC", "Peripheral - Barcode Reader" },
                new string[] { "PC-LAPTOP", "Computer - Laptop" },
                new string[] { "PC-", "Computer - Desktop" },
                new string[] { "MACBOOK", "Computer - Laptop" },
                new string[] { "IMAC", "Computer - Desktop" },
                new string[] { "MAC-", "Computer - Desktop" },
                new string[] { "SCANNER", "Scanner - ???" },
                new string[] { "SCANNER", "Scanner - ???" },
                new string[] { "SCAN-FUJITSU-IX500", "Scanner - ADF" },
                new string[] { "SCAN", "Scanner - ???" },
                new string[] { "SWITCH", "Networking - Switch - ???" },
            };

            public String[][] Model_Names = new string[][] {
                new string[] { "OPTIPLEX", "Dell Optiplex ???" },
                new string[] { "VOSTRO 220", "Dell Vostro 220" },
                new string[] { "VOSTRO", "Dell Vostro ???" },
                new string[] { "E1910HC", "Dell E1910Hc" },
                new string[] { "E1910H", "Dell E1910H" },
                new string[] { "P2211HT", "Dell P2211Ht" },
                new string[] { "P2211H", "Dell P2211H" },
                new string[] { "E228WFPF", "Dell E228WFPf" },
                new string[] { "E228WFP", "Dell E228WFP" },
                new string[] { "E177FPB", "Dell E177FPb" },
                new string[] { "E177FP", "Dell E177FP" },
                new string[] { "E178WFPF", "Dell E178WFPf" },
                new string[] { "E178WFP", "Dell E178WFP" },
                new string[] { "DELL", "Dell ???" },
                new string[] { "P4015N", "HP LaserJet P4015n" },
                new string[] { "P4015DN", "HP LaserJet P4015dn" },
                new string[] { "HP", "HP ???" },
                new string[] { "HEWLETT PACKARD", "HP ???" },
                new string[] { "STARTECH", "StarTech ???" },
                new string[] { "SCAN-FUJITSU-IX500", "Fujitsu Scan Snap iX500" },
                new string[] { "MACBOOK-PRO", "Apple Macbook Pro ???" },
                new string[] { "WORKCENTRE", "Xerox WorkCentre ???" }
            };

            public String[][] Manufacturers = new string[][] {
                new string[] { "XEROX", "Xerox" },
                new string[] { "LINKSYS", "Linksys" },
                new string[] { "AXIS", "Axis Communications" },
                new string[] { "MACBOOK-PRO", "Apple" },
                new string[] { "APC", "APC by Schneider Electric" },
                new string[] { "ACER", "Acer" },
                new string[] { "FUJITSU", "Fujitsu" },
                new string[] { "FUJI", "Fujitsu" },
                new string[] { "IMAC", "Apple" },
                new string[] { "HP", "Hewlett Packard" },
                new string[] { "HEWLETT PACKARD", "Hewlett Packard" },
                new string[] { "STARTECH", "StarTech" },
                new string[] { "DELL", "Dell" },
                new string[] { "OPTIPLEX", "Dell" },
                new string[] { "BELKIN", "Belkin" },
                new string[] { "EPSON", "Epson" },
                new string[] { "SEAGATE", "Seagate" },
                new string[] { "CISCO", "Cisco" },
                new string[] { "NETGEAR", "Netgear" },
                new string[] { "FARONICS", "Faronics" },
                new string[] { "SHARP", "Sharp" },
                new string[] { "SOLARWINDS", "SolarWinds" }
            };
            
            public MicroRecord_Translated(OleDbDataReader reader)
            {
                Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", reader["PRODUCT_NO"], reader["LOCATION"], reader["SERIAL_NUM"], reader["COMMENTS"]));

                String comments = reader["COMMENTS"].ToString();
                String serial = reader["SERIAL_NUM"].ToString();

                Boolean catalog_referenced = false;
                foreach (var v in item_catalog)
                {
                    if (comments.ToUpper().Contains(v.Key.ToUpper()) || serial.ToUpper().Contains(v.Key.ToUpper()))
                    {
                        Category = v.Value["Category"].ToString();
                        Manufacturer = v.Value["Manufacturer"].ToString();
                        Model_Name = v.Value["Model Name"].ToString();
                        Model_Number = v.Value["Model Number"].ToString();
                        Notes = v.Value["Notes"].ToString();

                        if (Category.Equals("Display - Monitor"))
                        {
                            if (v.Value["Display Size"] != null)
                            {
                                Display_Size = v.Value["Display Size"].ToString();
                            }
                            if (v.Value["Touchscreen"] != null)
                            {
                                Touchscreen = v.Value["Touchscreen"].ToString();
                            }
                        }

                        //mark that the catalog was referenced
                        catalog_referenced = true;
                    }
                }
                //else if the catalog was not referenced
                if (catalog_referenced == false)
                {
                    //check categories
                    foreach (string[] cate in Categories)
                    {
                        if (serial.ToUpper().Contains(cate[0]))
                        {
                            Category = cate[1];
                            break;
                        }
                    }
                    if (Category == null)
                    {
                        foreach (string[] cate in Categories)
                        {
                            if (comments.ToUpper().Contains(cate[0]))
                            {
                                Category = cate[1];
                                break;
                            }
                        }
                    }

                    //check manufacturers
                    foreach (string[] manu in Manufacturers)
                    {
                        if (comments.ToUpper().Contains(manu[0]))
                        {
                            Manufacturer = manu[1];
                            break;
                        }
                    }
                    if (Manufacturer == null)
                    {
                        foreach (string[] manu in Manufacturers)
                        {
                            if (serial.ToUpper().Contains(manu[0]))
                            {
                                Manufacturer = manu[1];
                                break;
                            }
                        }
                    }

                    //check model names
                    foreach (string[] modn in Model_Names)
                    {
                        if (comments.ToUpper().Contains(modn[0]))
                        {
                            Model_Name = modn[1];
                            break;
                        }
                    }
                    if (Model_Name == null)
                    {
                        foreach (string[] modn in Model_Names)
                        {
                            if (serial.ToUpper().Contains(modn[0]))
                            {
                                Model_Name = modn[1];
                                break;
                            }
                        }
                    }
                }

                ///////////////////////////////////////////////////////////////////////////

                //generate location
                string old_loc = reader["LOCATION"].ToString();

                string[] old_loc_split = old_loc.Split('-');

                if (old_loc_split.Length == 0)
                {
                    
                } else if (old_loc_split.Length == 1)
                {
                    Dept_Room_Location = old_loc_split[0];
                } else
                {
                    Name = String.Join("", (old_loc_split.Skip(1).ToArray()));
                    Dept_Room_Location = old_loc_split[0];
                }

                //set department location from old dept/room format
                foreach (string[] loca in SETTINGS.Location_Names)
                {
                    if (old_loc.ToUpper().Contains(loca[0]))
                    {
                        Location = loca[1];
                        break;
                    }
                }

                //generate primary serial
                Order_Number = reader["PO_NO"].ToString();

                if (reader["MONITOR"].ToString() != null && !reader["MONITOR"].ToString().Trim().Equals(""))
                {
                    Serial_Number = reader["MONITOR"].ToString();
                }
                //else if ((Manufacturer == null))
                else
                {
                    String serial_string = reader["SERIAL_NUM"].ToString();
                    if (serial_string.Contains('-'))
                    {
                        Serial_Number = serial_string.Substring(serial_string.LastIndexOf('-') + 1);
                        Console.WriteLine(serial_string);
                    }
                }

                //generate deployment status
                if (!Dept_Room_Location.Contains("DNS/110") && !Dept_Room_Location.Contains("DNS/111") && !Dept_Room_Location.Contains("DNS/017"))
                {
                    Status = "Deployed";
                }

                //generate note
                var multiline = string.Concat(
                    "[ Lock Combos ]", 
                    Environment.NewLine, 
                    reader["LOCK_COMBO"].ToString(), 
                    Environment.NewLine,
                    Environment.NewLine,
                    "[ Serial / Monitor ]",
                    Environment.NewLine,
                    reader["SERIAL_NUM"].ToString(),
                    " / ",
                    reader["MONITOR"].ToString(),
                    Environment.NewLine,
                    Environment.NewLine,
                    "[ Locations ]",
                    Environment.NewLine,
                    reader["LOCATION"].ToString(),
                    " <-- ",
                    reader["PREV_LOCAT"].ToString(),
                    Environment.NewLine,
                    Environment.NewLine,
                    "[ Comments ]",
                    Environment.NewLine,
                    reader["COMMENTS"].ToString());

                Notes = multiline;

                //Notes = String.Format("Lock Combos\r\n{0}\n\n{1}\n{2}\n{3}\n{4}\n{5}\n",
                //    reader["LOCK_COMBO"].ToString(),
                //    reader["PRODUCT_NO"].ToString(),
                //    reader["PO_NO"].ToString(),
                //    reader["SERIAL_NUM"].ToString(),
                //    reader["MONITOR"].ToString(),
                //    reader["COMMENTS"].ToString()
                //    );

                //copypasta'd from:
                //https://stackoverflow.com/questions/13584519/format-a-mac-address-using-string-format-in-c-sharp
                //format mac address
                var macaddress = reader["MAC_ADDRESS"].ToString().ToUpper();

                var regex = "(.{2})(.{2})(.{2})(.{2})(.{2})(.{2})";
                var replace = "$1:$2:$3:$4:$5:$6";
                Mac_Address_1 = Regex.Replace(macaddress, regex, replace);

                //generate service tag
                //if ((Manufacturer != null && Manufacturer.Equals("Dell")) && (Category != null && Category.Contains("Computer -")))
                //{
                //    String service_tag_string = reader["SERIAL_NUM"].ToString();
                //    Service_Tag = service_tag_string.Substring(service_tag_string.LastIndexOf('-') + 1);
                //}

                //assign asset tag from product no
                Asset_Tag = reader["PRODUCT_NO"].ToString();

                Decal = reader["DECALNUM"].ToString();

                RAM = reader["RAM"].ToString();

                Console.WriteLine(String.Format("\tCategory = {0}", Category));
                Console.WriteLine(String.Format("\tModel Name = {0}", Model_Name));
                Console.WriteLine(String.Format("\tManufacturer = {0}", Manufacturer));
                Console.WriteLine(String.Format("\tAsset Tag = {0}", Asset_Tag));
                Console.WriteLine(String.Format("\tDept/Room Location = {0}", Dept_Room_Location));
                Console.WriteLine();
            }
        }

    }
}
