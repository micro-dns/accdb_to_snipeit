using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace accdb_export
{
    class Settings
    {
        //Path to .accdb file
        public string DEFAULT_PATH { get; set; }

        public string DEFAULT_ITEM_INDEX { get; set; }

        //Query Parameters
        public String LIKE;
        public String PRODUCT_NUMBER;

        //Export Settings
        public String XLSX_PATH;
        public String CSV_PATH;
        public String EXPORT_TYPE; //xlsx, csv, etc

        //other fields
        public String COMPANY;

        //Location Names
        public String[][] Location_Names = new string[][] {
            new string[] { "MAPS/", "GOVDOCS - Government Documents" },
            new string[] { "ASIA/", "ASIA - Asia" },
            new string[] { "PNR/", "PNR - Preservations" },
            new string[] { "PRESERV/", "PNR - Preservations" },
            new string[] { "CAT/", "CAT - Cataloging" },
            new string[] { "CIRC/", "CIRC - Circulations" },
            new string[] { "SPEC COLL/", "HAWNPAC - Hawaiian and Pacific Special Collections" },
            new string[] { "SPECCOLL/", "HAWNPAC - Hawaiian and Pacific Special Collections" },
            new string[] { "HAWN PAC/", "HAWNPAC - Hawaiian and Pacific Special Collections" },
            new string[] { "HAWNPAC/", "HAWNPAC - Hawaiian and Pacific Special Collections" },
            new string[] { "CAT/027", "CLASSROOM - 027" },
            new string[] { "CLASSROOM/113", "CLASSROOM - 113" },
            new string[] { "CLASSROOM/156", "CLASSROOM - 156" },
            new string[] { "CLASSROOM/301", "CLASSROOM - 301" },
            new string[] { "CLASSROOM/306", "CLASSROOM - 306" },
            new string[] { "CLASSROOM/401", "CLASSROOM - 401" },
            new string[] { "DIGITIZING/", "DIGITIZING - Digitizing Lab" },
            new string[] { "FISCAL/", "FISCAL - Fiscal" },
            new string[] { "ACQUISITIONS/", "ACQ - Acquisitions" },
            new string[] { "ACQ/", "ACQ - Acquisitions" },
            new string[] { "MAILROOM/", "MAILROOM - Mailroom" },
            new string[] { "MAIL ROOM/", "MAILROOM - Mailroom" },
            new string[] { "MAIL/", "MAILROOM - Mailroom" },
            new string[] { "ARCHIVES/", "ARCH - Archives" },
            new string[] { "ARCH/", "ARCH - Archives" },
            new string[] { "MICROFILM/", "MICROFILM - Microfilm Reading Room" },
            new string[] { "MICROFORM/", "MICROFILM - Microfilm Reading Room" },
            new string[] { "BHSD/", "BHSD - Business, Humanities, and Social Sciences" },
            new string[] { "SCI TECH/", "SCITECH - Science and Technology" },
            new string[] { "SCITECH/", "SCITECH - Science and Technology" },
            new string[] { "CHARLOT/", "CHARLOT - Jean Charlot Collections" },
            new string[] { "SERIALS/", "SERIALS - Serials" },
            new string[] { "SYSTEMS/", "SYS - Systems" },
            new string[] { "SYS/", "SYS - Systems" },
            new string[] { "CUST/", "CUST - Custodian" },
            new string[] { "ESP/", "ESP - External Services Program" },
            new string[] { "GOVDOCS/", "GOVDOCS - Government Documents" },
            new string[] { "GOV DOCS/", "GOVDOCS - Government Documents" },
            new string[] { "ADMIN/", "ADMIN - Administration" },
        };

        public Settings()
        {
            //load settings
            //if settings file exists

            //if settings file does not exist
        }

        public void Save()
        {
            //save settings
        }
    }
}
