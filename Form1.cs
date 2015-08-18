using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;

namespace ProductRepository_Migration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnMigrate_Click(object sender, EventArgs e)
        {
            UpdateChargeAccounts();
            //MigrateV2();
        }

        public void MigrateV1()
        {
            string url = "http://spmaindev.volpe.dot.gov";
            //string url = "http://zebaduag03644"; 

            //string sub = "*";
            string sub = "/sites/Tools/Communications/ProjectRepository";

            SPSecurity.RunWithElevatedPrivileges(delegate
            {

                using (SPSite site = new SPSite(url + sub))
                {

                    //just forcing that the site opens
                    string temp = site.Owner.ToString();

                    //using (SPWeb web = (sub == "*") ? site.RootWeb : site.OpenWeb(sub))
                    //using (SPWeb web = site.OpenWeb(sub))
                    using (SPWeb web = site.OpenWeb())
                    {
                        //just forcing that the web opens
                        web.Site.OpenWeb(sub);
                        bool tempForWeb = web.Exists;
                        bool tempForWeb2 = web.UserIsSiteAdmin;

                        try
                        {
                            //Open library
                            string fullLibraryUrl = web.Url + "/Lists/DocumentLibrary/";
                            SPList list = (SPList)web.GetList(fullLibraryUrl);
                            SPListItemCollection listItems = list.Items;

                            //get data from database
                            string PRQuery =
                            "SELECT " +
                            "DISTINCT [DOC_ID], [TITLE], [SUBSITE_NAME], [AUTH_NAME], [ORG_NAME], [DocumentType], [YEAR], [NTIS_REPORT_NUM], [VOLPE_REPORT_NUM], [DOT_NUM], [ALT_REPORT_NUM], " +
                            "[REPORT_DATE], [VOLUME], [NUM_PAGES], [URL_1], [URL_2], [STATUS], [FILE_NAME], [TOPIC_NAME], [ABSTRACT], [NOTES], [Visibility], [KEYWORD1], [KEYWORD2], " +
                            "[KEYWORD3], [KEYWORD4], [KEYWORD5], [KEYWORD6], [KEYWORD7], [KEYWORD8], [KEYWORD9], [KEYWORD10], [KEYWORD11], [KEYWORD12] " +
                            "FROM " +
                            "( " +
                            "SELECT " +
                            "d.[DOC_ID], " +
                            "d.[TITLE], " +
                            "ISNULL(su.[SUBSITE_NAME], '') AS 'SUBSITE_NAME', " +
                            "ISNULL(d.[AUTH_NAME], '') AS 'AUTH_NAME', " +
                            "ISNULL(so.[ORG_NAME], '') AS 'ORG_NAME', " +
                            "ISNULL(dt.[TYPE_NAME], 'Report') as 'DocumentType', " +
                            "ISNULL(d.[YEAR], '0') AS 'YEAR', " +
                            "ISNULL(rd.[NTIS_REPORT_NUM], '') AS 'NTIS_REPORT_NUM', " +
                            "ISNULL(rd.[VOLPE_REPORT_NUM], '') AS 'VOLPE_REPORT_NUM', " +
                            "ISNULL(rd.[DOT_NUM], '') AS 'DOT_NUM', " +
                            "ISNULL(rd.[ALT_REPORT_NUM], '') AS 'ALT_REPORT_NUM', " +
                            "ISNULL(rd.[REPORT_DATE], '') AS 'REPORT_DATE', " +
                            "ISNULL(jd.[VOLUME], '') AS 'VOLUME', " +
                            "ISNULL(d.[NUM_PAGES], 0) AS 'NUM_PAGES', " +
                            "ISNULL(d.[URL], '') AS 'URL_1', " +
                            "ISNULL(d.[URL_2], '') AS 'URL_2', " +
                            "ISNULL(d.[STATUS], '') AS 'STATUS', " +
                            "ISNULL(d.[FILE_NAME], '') AS 'FILE_NAME', " +
                            "ISNULL(t.[TOPIC_NAME], '') AS 'TOPIC_NAME', " +
                            "ISNULL(d.[ABSTRACT], '') AS 'ABSTRACT', " +
                            "ISNULL(d.[NOTES], '') AS 'NOTES', " +
                            "'Live' AS 'Visibility', " +
                            "ISNULL(kt.[KEYWORD1], '') AS 'KEYWORD1', " +
                            "ISNULL(kt.[KEYWORD2], '') AS 'KEYWORD2', " +
                            "ISNULL(kt.[KEYWORD3], '') AS 'KEYWORD3', " +
                            "ISNULL(kt.[KEYWORD4], '') AS 'KEYWORD4', " +
                            "ISNULL(kt.[KEYWORD5], '') AS 'KEYWORD5', " +
                            "ISNULL(kt.[KEYWORD6], '') AS 'KEYWORD6', " +
                            "ISNULL(kt.[KEYWORD7], '') AS 'KEYWORD7', " +
                            "ISNULL(kt.[KEYWORD8], '') AS 'KEYWORD8', " +
                            "ISNULL(kt.[KEYWORD9], '') AS 'KEYWORD9', " +
                            "ISNULL(kt.[KEYWORD10], '') AS 'KEYWORD10', " +
                            "ISNULL(kt.[KEYWORD11], '') AS 'KEYWORD11', " +
                            "ISNULL(kt.[KEYWORD12], '') AS 'KEYWORD12' " +
                            "FROM [product_repository].[EARCHIVE].[DOCUMENTS_MAIN] d " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[SUBSITE] su ON d.SUBSITE_ID = su.SUBSITE_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[TOPICS] t ON d.TOPIC_ID = t.TOPIC_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[DOC_TYPE] dt ON d.DOC_TYPE = dt.[TYPE_ID] " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[SPONSOR_ORG_DOC] sod ON D.DOC_ID = SOD.DOC_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[SPONSOR_ORG] so ON sod.ORG_ID = so.ORG_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[JOURNALS_DETAIL] jd ON d.DOC_ID = jd.DOC_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[REPORTS_DETAIL] rd ON d.DOC_ID = rd.DOC_ID " +
                            "LEFT JOIN [product_repository].[EARCHIVE].[KEYWORDS_TRANSPOSED] kt ON d.DOC_ID = kt.DOC_ID " +
                            ") tmpA ORDER BY 1";

                            DataSet PRData = getQuery(PRQuery);

                            DataTable dataTable = new DataTable();
                            dataTable = PRData.Tables[0];

                            foreach (DataRow dataRow in dataTable.Rows)
                            {

                                //Adding a new item involves first "adding" and then updating it with the new values
                                //Column names are SP internal column names
                                SPListItem item = listItems.Add();

                                item["Title"] = dataRow["TITLE"].ToString();

                                SPList subsitesList = web.Lists["Subsites"];
                                SPQuery querySubsites = new SPQuery();
                                int intSelectedIdSubsites = 0;
                                querySubsites.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["SUBSITE_NAME"].ToString() + "</Value></Eq></Where>";
                                if (subsitesList.GetItems(querySubsites) != null)
                                {
                                    try
                                    {
                                        SPListItem result = subsitesList.GetItems(querySubsites)[0];
                                        intSelectedIdSubsites = result.ID;
                                    }
                                    catch (Exception esu)
                                    {
                                        intSelectedIdSubsites = 0;
                                        esu.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdSubsites = 0;
                                }

                                if (intSelectedIdSubsites > 0)
                                {
                                    item["Subsite"] = new SPFieldLookupValue(intSelectedIdSubsites, dataRow["SUBSITE_NAME"].ToString());
                                }
                                else
                                {
                                    //item["Subsite"] = new SPFieldLookupValue(1, "None");
                                }

                                item["Author0"] = dataRow["AUTH_NAME"].ToString();

                                SPList sponsorsList = web.Lists["Sponsors"];
                                SPQuery querySponsors = new SPQuery();
                                int intSelectedIdSponsors = 0;
                                querySponsors.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["ORG_NAME"].ToString() + "</Value></Eq></Where>";
                                if (sponsorsList.GetItems(querySponsors) != null)
                                {
                                    try
                                    {
                                        SPListItem result = sponsorsList.GetItems(querySponsors)[0];
                                        intSelectedIdSponsors = result.ID;
                                    }
                                    catch (Exception esp)
                                    {
                                        intSelectedIdSponsors = 0;
                                        esp.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdSponsors = 0;
                                }

                                if (intSelectedIdSponsors > 0)
                                {
                                    item["Sponsor"] = new SPFieldLookupValue(intSelectedIdSponsors, dataRow["ORG_NAME"].ToString());
                                }
                                else
                                {
                                    //item["Sponsor"] = new SPFieldLookupValue(1, "None");
                                }

                                item["DocumentType"] = dataRow["DocumentType"].ToString();

                                item["Year"] = dataRow["Year"].ToString();
                                item["NTISReportNumber"] = dataRow["NTIS_REPORT_NUM"].ToString();
                                item["VolpeReportNumber"] = dataRow["VOLPE_REPORT_NUM"].ToString();
                                item["DOTNumber"] = dataRow["DOT_NUM"].ToString();
                                item["AlternateReportNumber"] = dataRow["ALT_REPORT_NUM"].ToString();

                                DateTime reportDate = (DateTime)dataRow["REPORT_DATE"];
                                item["DatePublished"] = reportDate;

                                item["Volume"] = dataRow["Volume"].ToString();
                                item["NumberOfPages"] = dataRow["NUM_PAGES"].ToString();

                                SPFieldUrlValue url1Value = new SPFieldUrlValue();
                                url1Value.Description = dataRow["URL_1"].ToString();
                                url1Value.Url = dataRow["URL_1"].ToString();
                                item["URL1"] = url1Value;

                                SPFieldUrlValue url2Value = new SPFieldUrlValue();
                                url2Value.Description = dataRow["URL_2"].ToString();
                                url2Value.Url = dataRow["URL_2"].ToString();
                                item["URL2"] = url2Value;

                                item["Status"] = dataRow["STATUS"].ToString();

                                item["FileName"] = dataRow["FILE_NAME"].ToString();


                                SPList topicsList = web.Lists["Topics"];
                                SPQuery queryTopics = new SPQuery();
                                int intSelectedIdTopics = 0;
                                queryTopics.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["TOPIC_NAME"].ToString() + "</Value></Eq></Where>";
                                if (topicsList.GetItems(queryTopics) != null)
                                {
                                    try
                                    {
                                        SPListItem result = topicsList.GetItems(queryTopics)[0];
                                        intSelectedIdTopics = result.ID;
                                    }
                                    catch (Exception eto)
                                    {
                                        intSelectedIdTopics = 0;
                                        eto.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdTopics = 0;
                                }

                                if (intSelectedIdTopics > 0)
                                {
                                    item["Topic"] = new SPFieldLookupValue(intSelectedIdTopics, dataRow["TOPIC_NAME"].ToString());
                                }
                                else
                                {
                                    //item["Topic"] = new SPFieldLookupValue(1, "None");
                                }


                                item["Abstract"] = dataRow["ABSTRACT"].ToString();
                                item["Notes"] = dataRow["NOTES"].ToString();

                                item["Visibility"] = dataRow["Visibility"].ToString();


                                SPList Keyword1List = web.Lists["Keywords"];
                                SPQuery queryKeyword1 = new SPQuery();
                                int intSelectedIdKeyword1 = 0;
                                queryKeyword1.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD1"].ToString() + "</Value></Eq></Where>";
                                if (Keyword1List.GetItems(queryKeyword1) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword1List.GetItems(queryKeyword1)[0];
                                        intSelectedIdKeyword1 = result.ID;
                                    }
                                    catch (Exception e1)
                                    {
                                        intSelectedIdKeyword1 = 0;
                                        e1.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword1 = 0;
                                }

                                if (intSelectedIdKeyword1 > 0)
                                {
                                    item["Keyword1"] = new SPFieldLookupValue(intSelectedIdKeyword1, dataRow["KEYWORD1"].ToString());
                                }
                                else
                                {
                                    //item["Keyword1"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword2

                                SPList Keyword2List = web.Lists["Keywords"];
                                SPQuery queryKeyword2 = new SPQuery();
                                int intSelectedIdKeyword2 = 0;
                                queryKeyword2.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD2"].ToString() + "</Value></Eq></Where>";
                                if (Keyword2List.GetItems(queryKeyword2) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword2List.GetItems(queryKeyword2)[0];
                                        intSelectedIdKeyword2 = result.ID;
                                    }
                                    catch (Exception e2)
                                    {
                                        intSelectedIdKeyword2 = 0;
                                        e2.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword2 = 0;
                                }

                                if (intSelectedIdKeyword2 > 0)
                                {
                                    item["Keyword2"] = new SPFieldLookupValue(intSelectedIdKeyword2, dataRow["KEYWORD2"].ToString());
                                }
                                else
                                {
                                    //item["Keyword2"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword3

                                SPList Keyword3List = web.Lists["Keywords"];
                                SPQuery queryKeyword3 = new SPQuery();
                                int intSelectedIdKeyword3 = 0;
                                queryKeyword3.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD3"].ToString() + "</Value></Eq></Where>";
                                if (Keyword3List.GetItems(queryKeyword3) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword3List.GetItems(queryKeyword3)[0];
                                        intSelectedIdKeyword3 = result.ID;
                                    }
                                    catch (Exception e3)
                                    {
                                        intSelectedIdKeyword3 = 0;
                                        e3.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword3 = 0;
                                }

                                if (intSelectedIdKeyword3 > 0)
                                {
                                    item["Keyword3"] = new SPFieldLookupValue(intSelectedIdKeyword3, dataRow["KEYWORD3"].ToString());
                                }
                                else
                                {
                                    //item["Keyword3"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword4

                                SPList Keyword4List = web.Lists["Keywords"];
                                SPQuery queryKeyword4 = new SPQuery();
                                int intSelectedIdKeyword4 = 0;
                                queryKeyword4.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD4"].ToString() + "</Value></Eq></Where>";
                                if (Keyword4List.GetItems(queryKeyword4) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword4List.GetItems(queryKeyword4)[0];
                                        intSelectedIdKeyword4 = result.ID;
                                    }
                                    catch (Exception e4)
                                    {
                                        intSelectedIdKeyword4 = 0;
                                        e4.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword4 = 0;
                                }

                                if (intSelectedIdKeyword4 > 0)
                                {
                                    item["Keyword4"] = new SPFieldLookupValue(intSelectedIdKeyword4, dataRow["KEYWORD4"].ToString());
                                }
                                else
                                {
                                    //item["Keyword4"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword5

                                SPList Keyword5List = web.Lists["Keywords"];
                                SPQuery queryKeyword5 = new SPQuery();
                                int intSelectedIdKeyword5 = 0;
                                queryKeyword5.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD5"].ToString() + "</Value></Eq></Where>";
                                if (Keyword5List.GetItems(queryKeyword5) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword5List.GetItems(queryKeyword5)[0];
                                        intSelectedIdKeyword5 = result.ID;
                                    }
                                    catch (Exception e5)
                                    {
                                        intSelectedIdKeyword5 = 0;
                                        e5.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword5 = 0;
                                }

                                if (intSelectedIdKeyword5 > 0)
                                {
                                    item["Keyword5"] = new SPFieldLookupValue(intSelectedIdKeyword5, dataRow["KEYWORD5"].ToString());
                                }
                                else
                                {
                                    //item["Keyword5"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword6

                                SPList Keyword6List = web.Lists["Keywords"];
                                SPQuery queryKeyword6 = new SPQuery();
                                int intSelectedIdKeyword6 = 0;
                                queryKeyword6.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD6"].ToString() + "</Value></Eq></Where>";
                                if (Keyword6List.GetItems(queryKeyword6) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword6List.GetItems(queryKeyword6)[0];
                                        intSelectedIdKeyword6 = result.ID;
                                    }
                                    catch (Exception e6)
                                    {
                                        intSelectedIdKeyword6 = 0;
                                        e6.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword6 = 0;
                                }

                                if (intSelectedIdKeyword6 > 0)
                                {
                                    item["Keyword6"] = new SPFieldLookupValue(intSelectedIdKeyword6, dataRow["KEYWORD6"].ToString());
                                }
                                else
                                {
                                    //item["Keyword6"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword7

                                SPList Keyword7List = web.Lists["Keywords"];
                                SPQuery queryKeyword7 = new SPQuery();
                                int intSelectedIdKeyword7 = 0;
                                queryKeyword7.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD7"].ToString() + "</Value></Eq></Where>";
                                if (Keyword7List.GetItems(queryKeyword7) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword7List.GetItems(queryKeyword7)[0];
                                        intSelectedIdKeyword7 = result.ID;
                                    }
                                    catch (Exception e7)
                                    {
                                        intSelectedIdKeyword7 = 0;
                                        e7.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword7 = 0;
                                }

                                if (intSelectedIdKeyword7 > 0)
                                {
                                    item["Keyword7"] = new SPFieldLookupValue(intSelectedIdKeyword7, dataRow["KEYWORD7"].ToString());
                                }
                                else
                                {
                                    //item["Keyword7"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword8

                                SPList Keyword8List = web.Lists["Keywords"];
                                SPQuery queryKeyword8 = new SPQuery();
                                int intSelectedIdKeyword8 = 0;
                                queryKeyword8.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD8"].ToString() + "</Value></Eq></Where>";
                                if (Keyword8List.GetItems(queryKeyword8) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword8List.GetItems(queryKeyword8)[0];
                                        intSelectedIdKeyword8 = result.ID;
                                    }
                                    catch (Exception e8)
                                    {
                                        intSelectedIdKeyword8 = 0;
                                        e8.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword8 = 0;
                                }

                                if (intSelectedIdKeyword8 > 0)
                                {
                                    item["Keyword8"] = new SPFieldLookupValue(intSelectedIdKeyword8, dataRow["KEYWORD8"].ToString());
                                }
                                else
                                {
                                    //item["Keyword8"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword9

                                SPList Keyword9List = web.Lists["Keywords"];
                                SPQuery queryKeyword9 = new SPQuery();
                                int intSelectedIdKeyword9 = 0;
                                queryKeyword9.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD9"].ToString() + "</Value></Eq></Where>";
                                if (Keyword9List.GetItems(queryKeyword9) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword9List.GetItems(queryKeyword9)[0];
                                        intSelectedIdKeyword9 = result.ID;
                                    }
                                    catch (Exception e9)
                                    {
                                        intSelectedIdKeyword9 = 0;
                                        e9.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword9 = 0;
                                }

                                if (intSelectedIdKeyword9 > 0)
                                {
                                    item["Keyword9"] = new SPFieldLookupValue(intSelectedIdKeyword9, dataRow["KEYWORD9"].ToString());
                                }
                                else
                                {
                                    //item["Keyword9"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword10

                                SPList Keyword10List = web.Lists["Keywords"];
                                SPQuery queryKeyword10 = new SPQuery();
                                int intSelectedIdKeyword10 = 0;
                                queryKeyword10.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD10"].ToString() + "</Value></Eq></Where>";
                                if (Keyword10List.GetItems(queryKeyword10) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword10List.GetItems(queryKeyword10)[0];
                                        intSelectedIdKeyword10 = result.ID;
                                    }
                                    catch (Exception e10)
                                    {
                                        intSelectedIdKeyword10 = 0;
                                        e10.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword10 = 0;
                                }

                                if (intSelectedIdKeyword10 > 0)
                                {
                                    item["Keyword10"] = new SPFieldLookupValue(intSelectedIdKeyword10, dataRow["KEYWORD10"].ToString());
                                }
                                else
                                {
                                    //item["Keyword10"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword11

                                SPList Keyword11List = web.Lists["Keywords"];
                                SPQuery queryKeyword11 = new SPQuery();
                                int intSelectedIdKeyword11 = 0;
                                queryKeyword11.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD11"].ToString() + "</Value></Eq></Where>";
                                if (Keyword11List.GetItems(queryKeyword11) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword11List.GetItems(queryKeyword11)[0];
                                        intSelectedIdKeyword11 = result.ID;
                                    }
                                    catch (Exception e11)
                                    {
                                        intSelectedIdKeyword11 = 0;
                                        e11.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword11 = 0;
                                }

                                if (intSelectedIdKeyword11 > 0)
                                {
                                    item["Keyword11"] = new SPFieldLookupValue(intSelectedIdKeyword11, dataRow["KEYWORD11"].ToString());
                                }
                                else
                                {
                                    //item["Keyword11"] = new SPFieldLookupValue(1, "None");
                                }

                                //Keyword12

                                SPList Keyword12List = web.Lists["Keywords"];
                                SPQuery queryKeyword12 = new SPQuery();
                                int intSelectedIdKeyword12 = 0;
                                queryKeyword12.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["KEYWORD12"].ToString() + "</Value></Eq></Where>";
                                if (Keyword12List.GetItems(queryKeyword12) != null)
                                {
                                    try
                                    {
                                        SPListItem result = Keyword12List.GetItems(queryKeyword12)[0];
                                        intSelectedIdKeyword12 = result.ID;
                                    }
                                    catch (Exception e12)
                                    {
                                        intSelectedIdKeyword12 = 0;
                                        e12.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdKeyword12 = 0;
                                }

                                if (intSelectedIdKeyword12 > 0)
                                {
                                    item["Keyword12"] = new SPFieldLookupValue(intSelectedIdKeyword12, dataRow["KEYWORD12"].ToString());
                                }
                                else
                                {
                                    //item["Keyword12"] = new SPFieldLookupValue(1, "None");
                                }

                                //Verify that strings are not larger than expected

                                if (item["Title"].ToString().Length > 255)
                                {
                                    item["Title"] = item["Title"].ToString().Substring(0, 255);
                                }

                                if (item["Author0"].ToString().Length > 255)
                                {
                                    item["Author0"] = item["Author0"].ToString().Substring(0, 255);
                                }

                                if (item["DocumentType"].ToString().Length > 255)
                                {
                                    item["DocumentType"] = item["DocumentType"].ToString().Substring(0, 255);
                                }

                                if (item["Year"].ToString().Length > 255)
                                {
                                    item["Year"] = item["Year"].ToString().Substring(0, 255);
                                }

                                if (item["NTISReportNumber"].ToString().Length > 255)
                                {
                                    item["NTISReportNumber"] = item["NTISReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["VolpeReportNumber"].ToString().Length > 255)
                                {
                                    item["VolpeReportNumber"] = item["VolpeReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["DOTNumber"].ToString().Length > 255)
                                {
                                    item["DOTNumber"] = item["DOTNumber"].ToString().Substring(0, 255);
                                }

                                if (item["AlternateReportNumber"].ToString().Length > 255)
                                {
                                    item["AlternateReportNumber"] = item["AlternateReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["Volume"].ToString().Length > 255)
                                {
                                    item["Volume"] = item["Volume"].ToString().Substring(0, 255);
                                }

                                if (item["NumberOfPages"].ToString().Length > 255)
                                {
                                    item["NumberOfPages"] = item["NumberOfPages"].ToString().Substring(0, 255);
                                }

                                if (item["Status"].ToString().Length > 255)
                                {
                                    item["Status"] = item["Status"].ToString().Substring(0, 255);
                                }

                                if (item["FileName"].ToString().Length > 255)
                                {
                                    item["FileName"] = item["FileName"].ToString().Substring(0, 255);
                                }

                                if (item["Visibility"].ToString().Length > 255)
                                {
                                    item["Visibility"] = item["Visibility"].ToString().Substring(0, 255);
                                }

                                if (item["Abstract"].ToString().Length > 255)
                                {
                                    item["Abstract"] = item["Abstract"].ToString().Substring(0, 255);
                                }

                                if (item["Notes"].ToString().Length > 255)
                                {
                                    item["Notes"] = item["Notes"].ToString().Substring(0, 255);
                                }

                                //Finally, save the item with the updated values
                                item.Update();

                            }

                            Console.WriteLine("Done. Refreshing...");
                            list.Update();

                            btnMigrate.Text = "Done. Migrate again?";

                        }
                        catch (Exception ex)
                        {
                            SPDiagnosticsService diagSvc = SPDiagnosticsService.Local;
                            diagSvc.WriteTrace(0, new SPDiagnosticsCategory("ProjectRepository", TraceSeverity.Monitorable, EventSeverity.Error),
                            TraceSeverity.Monitorable, "ProjectRepository error:  {0}", new object[] { ex.ToString() });
                        }

                    }
                }

            });
        }

        public void MigrateV2() 
        {
            string url = "http://spmaindev.volpe.dot.gov";
            //string url = "http://zebaduag03644"; 
            
            //string sub = "*";
            string sub = "/sites/Tools/Communications/ProjectRepository";

            SPSecurity.RunWithElevatedPrivileges(delegate
            {

                using (SPSite site = new SPSite(url + sub))
                {
                    
                    //just forcing that the site opens
                    string temp = site.Owner.ToString();
                    
                    //using (SPWeb web = (sub == "*") ? site.RootWeb : site.OpenWeb(sub))
                    //using (SPWeb web = site.OpenWeb(sub))
                    using (SPWeb web = site.OpenWeb())
                    {
                        //just forcing that the web opens
                        web.Site.OpenWeb(sub);
                        bool tempForWeb = web.Exists;
                        bool tempForWeb2 = web.UserIsSiteAdmin;

                        try
                        {
                            //Open library
                            string fullLibraryUrl = web.Url + "/Lists/DocumentLibrary/";
                            SPList list = (SPList) web.GetList(fullLibraryUrl);
                            SPListItemCollection listItems = list.Items;
                    
                            //get data from database
                            string PRQuery =
                            "SELECT " +
                            "DISTINCT [DOC_ID], [TITLE], [SUBSITE_NAME], [AUTH_NAME], [ORG_NAME], [DocumentType], [YEAR], [NTIS_REPORT_NUM], [VOLPE_REPORT_NUM], [DOT_NUM], [ALT_REPORT_NUM], " +
                            "[REPORT_DATE], [VOLUME], [NUM_PAGES], [URL_1], [URL_2], [STATUS], [ABSTRACT], [NOTES], [Visibility], [KEYWORD1], [KEYWORD2], " +
                            "[KEYWORD3], [KEYWORD4], [KEYWORD5], [KEYWORD6], [KEYWORD7], [KEYWORD8], [KEYWORD9], [KEYWORD10], [KEYWORD11], [KEYWORD12], " +
                            "[IAAProjectNumber1], [IAAProjectNumber2], [IAAProjectNumber3], [IAAProjectNumber4], [IAAProjectNumber5] " +
                            "FROM " +
                            "( " +
                            "SELECT " +
                            "d.[ID] AS 'DOC_ID', " +
                            "d.[TITLE], " +
                            "ISNULL(su.[Subsite_Name], '') AS 'SUBSITE_NAME', " +
                            "ISNULL(d.[Author], '') AS 'AUTH_NAME', " +
                            "ISNULL(so.[Name], '') AS 'ORG_NAME', " +
                            "ISNULL(dt.[TYPE_NAME], 'Report') as 'DocumentType', " +
                            "ISNULL(d.[YEAR], '0') AS 'YEAR', " +
                            "ISNULL(rd.[NTIS_Number], '') AS 'NTIS_REPORT_NUM', " +
                            "ISNULL(rd.[Volpe_Number], '') AS 'VOLPE_REPORT_NUM', " +
                            "ISNULL(rd.[DOT_Number], '') AS 'DOT_NUM', " +
                            "ISNULL(rd.[Alternate_Number], '') AS 'ALT_REPORT_NUM', " +
                            "ISNULL(rd.[Report_Date], '') AS 'REPORT_DATE', " +
                            "ISNULL(jd.[VOLUME], '') AS 'VOLUME', " +
                            "ISNULL(d.[Number_of_Pages], 0) AS 'NUM_PAGES', " +
                            "ISNULL(d.[URL], '') AS 'URL_1', " +
                            "ISNULL(d.[URL_2], '') AS 'URL_2', " +
                            "ISNULL(d.[STATUS], '') AS 'STATUS', " +
                            "ISNULL(d.[ABSTRACT], '') AS 'ABSTRACT', " +
                            "ISNULL(d.[NOTES], '') AS 'NOTES', " +
                            "'Live' AS 'Visibility', " +
                            "ISNULL(kt.[KEYWORD1], '') AS 'KEYWORD1', " +
                            "ISNULL(kt.[KEYWORD2], '') AS 'KEYWORD2', " +
                            "ISNULL(kt.[KEYWORD3], '') AS 'KEYWORD3', " +
                            "ISNULL(kt.[KEYWORD4], '') AS 'KEYWORD4', " +
                            "ISNULL(kt.[KEYWORD5], '') AS 'KEYWORD5', " +
                            "ISNULL(kt.[KEYWORD6], '') AS 'KEYWORD6', " +
                            "ISNULL(kt.[KEYWORD7], '') AS 'KEYWORD7', " +
                            "ISNULL(kt.[KEYWORD8], '') AS 'KEYWORD8', " +
                            "ISNULL(kt.[KEYWORD9], '') AS 'KEYWORD9', " +
                            "ISNULL(kt.[KEYWORD10], '') AS 'KEYWORD10', " +
                            "ISNULL(kt.[KEYWORD11], '') AS 'KEYWORD11', " +
                            "ISNULL(kt.[KEYWORD12], '') AS 'KEYWORD12', " +
                            "ISNULL([IAA_Project_Number_1], '') AS 'IAAProjectNumber1', " +
                            "ISNULL([IAA_Project_Number_2], '') AS 'IAAProjectNumber2', " +
                            "ISNULL([IAA_Project_Number_3], '') AS 'IAAProjectNumber3', " +
                            "ISNULL([IAA_Project_Number_4], '') AS 'IAAProjectNumber4', " +
                            "'' AS 'IAAProjectNumber5' " +
                            "FROM [Staging].[dbo].[Project_Document] d " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Site] su ON d.SITE_ID = su.ID " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Type] dt ON d.[Type_Id] = dt.[ID] " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Sponsor_JOIN] sod ON D.ID = SOD.[Document_Id] " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Sponsor] so ON sod.[Sponsor_Id] = so.ID " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Journal_Detail] jd ON d.ID = jd.Document_Id " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Report_Detail] rd ON d.ID = rd.Document_Id " +
                            "LEFT JOIN [Staging].[dbo].[Project_Document_Keywords_Transposed] kt ON d.ID = kt.DOC_ID " +
                            ") tmpA ORDER BY 1";
                        
                            DataSet PRData = getQuery(PRQuery);

                            DataTable dataTable = new DataTable();
                            dataTable = PRData.Tables[0];

                            foreach (DataRow dataRow in dataTable.Rows)
                            {

                                //Adding a new item involves first "adding" and then updating it with the new values
                                //Column names are SP internal column names
                                SPListItem item = listItems.Add();

                                item["Title"] = dataRow["TITLE"].ToString();

                                SPList subsitesList = web.Lists["Subsites"];
                                SPQuery querySubsites = new SPQuery();
                                int intSelectedIdSubsites = 0;
                                querySubsites.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["SUBSITE_NAME"].ToString() + "</Value></Eq></Where>";
                                if (subsitesList.GetItems(querySubsites) != null)
                                {
                                    try
                                    {
                                        SPListItem result = subsitesList.GetItems(querySubsites)[0];
                                        intSelectedIdSubsites = result.ID;
                                    }
                                    catch (Exception esu)
                                    {
                                        intSelectedIdSubsites = 0;
                                        esu.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdSubsites = 0;
                                }

                                if (intSelectedIdSubsites > 0)
                                {
                                    item["Subsite"] = new SPFieldLookupValue(intSelectedIdSubsites, dataRow["SUBSITE_NAME"].ToString());
                                }
                                else
                                {
                                    //item["Subsite"] = new SPFieldLookupValue(1, "None");
                                }

                                item["Author0"] = dataRow["AUTH_NAME"].ToString();

                                SPList sponsorsList = web.Lists["Sponsors"];
                                SPQuery querySponsors = new SPQuery();
                                int intSelectedIdSponsors = 0;
                                querySponsors.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["ORG_NAME"].ToString() + "</Value></Eq></Where>";
                                if (sponsorsList.GetItems(querySponsors) != null)
                                {
                                    try
                                    {
                                        SPListItem result = sponsorsList.GetItems(querySponsors)[0];
                                        intSelectedIdSponsors = result.ID;
                                    }
                                    catch (Exception esp)
                                    {
                                        intSelectedIdSponsors = 0;
                                        esp.ToString();
                                    }
                                }
                                else
                                {
                                    intSelectedIdSponsors = 0;
                                }

                                if (intSelectedIdSponsors > 0)
                                {
                                    item["Sponsor"] = new SPFieldLookupValue(intSelectedIdSponsors, dataRow["ORG_NAME"].ToString());
                                }
                                else
                                {
                                    //item["Sponsor"] = new SPFieldLookupValue(1, "None");
                                }

                                item["DocumentType"] = dataRow["DocumentType"].ToString();

                                item["Year"] = dataRow["Year"].ToString();
                                item["NTISReportNumber"] = dataRow["NTIS_REPORT_NUM"].ToString();
                                item["VolpeReportNumber"] = dataRow["VOLPE_REPORT_NUM"].ToString();
                                item["DOTNumber"] = dataRow["DOT_NUM"].ToString();
                                item["AlternateReportNumber"] = dataRow["ALT_REPORT_NUM"].ToString();

                                DateTime reportDate = (DateTime)dataRow["REPORT_DATE"];
                                item["DatePublished"] = reportDate;

                                item["Volume"] = dataRow["Volume"].ToString();
                                item["NumberOfPages"] = dataRow["NUM_PAGES"].ToString();

                                SPFieldUrlValue url1Value = new SPFieldUrlValue();
                                url1Value.Description = dataRow["URL_1"].ToString();
                                url1Value.Url = dataRow["URL_1"].ToString();
                                item["URL1"] = url1Value;

                                SPFieldUrlValue url2Value = new SPFieldUrlValue();
                                url2Value.Description = dataRow["URL_2"].ToString();
                                url2Value.Url = dataRow["URL_2"].ToString();
                                item["URL2"] = url2Value;

                                item["Status"] = dataRow["STATUS"].ToString();

                                item["Abstract"] = dataRow["ABSTRACT"].ToString();
                                item["Notes"] = dataRow["NOTES"].ToString();

                                item["Visibility"] = dataRow["Visibility"].ToString();

                                item["Keyword1"] = dataRow["KEYWORD1"].ToString();
                                item["Keyword2"] = dataRow["KEYWORD2"].ToString();
                                item["Keyword3"] = dataRow["KEYWORD3"].ToString();
                                item["Keyword4"] = dataRow["KEYWORD4"].ToString();
                                item["Keyword5"] = dataRow["KEYWORD5"].ToString();
                                item["Keyword6"] = dataRow["KEYWORD6"].ToString();
                                item["Keyword7"] = dataRow["KEYWORD7"].ToString();
                                item["Keyword8"] = dataRow["KEYWORD8"].ToString();
                                item["Keyword9"] = dataRow["KEYWORD9"].ToString();
                                item["Keyword10"] = dataRow["KEYWORD10"].ToString();
                                item["Keyword11"] = dataRow["KEYWORD11"].ToString();
                                item["Keyword12"] = dataRow["KEYWORD12"].ToString();

                                item["ChargeAccount1"] = dataRow["IAAProjectNumber1"].ToString();
                                item["ChargeAccount2"] = dataRow["IAAProjectNumber2"].ToString();
                                item["ChargeAccount3"] = dataRow["IAAProjectNumber3"].ToString();
                                item["ChargeAccount4"] = dataRow["IAAProjectNumber4"].ToString();
                                item["ChargeAccount5"] = dataRow["IAAProjectNumber5"].ToString();
                                
                                //Verify that strings are not larger than expected

                                if (item["Title"].ToString().Length > 255)
                                {
                                    item["Title"] = item["Title"].ToString().Substring(0, 255);
                                }

                                if (item["Author0"].ToString().Length > 255)
                                {
                                    item["Author0"] = item["Author0"].ToString().Substring(0, 255);
                                }

                                if (item["DocumentType"].ToString().Length > 255)
                                {
                                    item["DocumentType"] = item["DocumentType"].ToString().Substring(0, 255);
                                }

                                if (item["Year"].ToString().Length > 255)
                                {
                                    item["Year"] = item["Year"].ToString().Substring(0, 255);
                                }

                                if (item["NTISReportNumber"].ToString().Length > 255)
                                {
                                    item["NTISReportNumber"] = item["NTISReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["VolpeReportNumber"].ToString().Length > 255)
                                {
                                    item["VolpeReportNumber"] = item["VolpeReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["DOTNumber"].ToString().Length > 255)
                                {
                                    item["DOTNumber"] = item["DOTNumber"].ToString().Substring(0, 255);
                                }

                                if (item["AlternateReportNumber"].ToString().Length > 255)
                                {
                                    item["AlternateReportNumber"] = item["AlternateReportNumber"].ToString().Substring(0, 255);
                                }

                                if (item["Volume"].ToString().Length > 255)
                                {
                                    item["Volume"] = item["Volume"].ToString().Substring(0, 255);
                                }

                                if (item["NumberOfPages"].ToString().Length > 255)
                                {
                                    item["NumberOfPages"] = item["NumberOfPages"].ToString().Substring(0, 255);
                                }

                                if (item["Status"].ToString().Length > 255)
                                {
                                    item["Status"] = item["Status"].ToString().Substring(0, 255);
                                }

                                if (item["Visibility"].ToString().Length > 255)
                                {
                                    item["Visibility"] = item["Visibility"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword1"].ToString().Length > 255)
                                {
                                    item["Keyword1"] = item["Keyword1"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword2"].ToString().Length > 255)
                                {
                                    item["Keyword2"] = item["Keyword2"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword3"].ToString().Length > 255)
                                {
                                    item["Keyword3"] = item["Keyword3"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword4"].ToString().Length > 255)
                                {
                                    item["Keyword4"] = item["Keyword4"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword5"].ToString().Length > 255)
                                {
                                    item["Keyword5"] = item["Keyword5"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword6"].ToString().Length > 255)
                                {
                                    item["Keyword6"] = item["Keyword6"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword7"].ToString().Length > 255)
                                {
                                    item["Keyword7"] = item["Keyword7"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword8"].ToString().Length > 255)
                                {
                                    item["Keyword8"] = item["Keyword8"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword9"].ToString().Length > 255)
                                {
                                    item["Keyword9"] = item["Keyword9"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword10"].ToString().Length > 255)
                                {
                                    item["Keyword10"] = item["Keyword10"].ToString().Substring(0, 255);
                                }

                                if (item["Keyword11"].ToString().Length > 255)
                                {
                                    item["Keyword11"] = item["Keyword11"].ToString().Substring(0, 255);
                                }

                                if (item["Abstract"].ToString().Length > 255)
                                {
                                    item["Abstract"] = item["Abstract"].ToString().Substring(0, 255);
                                }

                                if (item["Notes"].ToString().Length > 255)
                                {
                                    item["Notes"] = item["Notes"].ToString().Substring(0, 255);
                                }

                                if (item["ChargeAccount1"].ToString().Length > 255)
                                {
                                    item["ChargeAccount1"] = item["ChargeAccount1"].ToString().Substring(0, 255);
                                }

                                if (item["ChargeAccount2"].ToString().Length > 255)
                                {
                                    item["ChargeAccount2"] = item["ChargeAccount2"].ToString().Substring(0, 255);
                                }

                                if (item["ChargeAccount3"].ToString().Length > 255)
                                {
                                    item["ChargeAccount3"] = item["ChargeAccount3"].ToString().Substring(0, 255);
                                }

                                if (item["ChargeAccount4"].ToString().Length > 255)
                                {
                                    item["ChargeAccount4"] = item["ChargeAccount4"].ToString().Substring(0, 255);
                                }

                                if (item["ChargeAccount5"].ToString().Length > 255)
                                {
                                    item["ChargeAccount5"] = item["ChargeAccount5"].ToString().Substring(0, 255);
                                }

                                //Finally, save the item with the updated values
                                item.Update();

                            }

                            Console.WriteLine("Done. Refreshing...");
                            list.Update();

                            btnMigrate.Text = "Done. Migrate again?";

                        }
                        catch (Exception ex)
                        {
                            SPDiagnosticsService diagSvc = SPDiagnosticsService.Local;
                            diagSvc.WriteTrace(0, new SPDiagnosticsCategory("ProjectRepository", TraceSeverity.Monitorable, EventSeverity.Error),
                            TraceSeverity.Monitorable, "ProjectRepository error:  {0}", new object[] { ex.ToString() });
                        }

                    }
                }

            });
        }

        public void UpdateChargeAccounts()
        {
            string url = "http://spmaindev.volpe.dot.gov";
            //string url = "http://zebaduag03644"; 

            //string sub = "*";
            string sub = "/sites/Tools/Communications/ProjectRepository";

            SPSecurity.RunWithElevatedPrivileges(delegate
            {

                using (SPSite site = new SPSite(url + sub))
                {

                    //just forcing that the site opens
                    string temp = site.Owner.ToString();

                    //using (SPWeb web = (sub == "*") ? site.RootWeb : site.OpenWeb(sub))
                    //using (SPWeb web = site.OpenWeb(sub))
                    using (SPWeb web = site.OpenWeb())
                    {
                        //just forcing that the web opens
                        web.Site.OpenWeb(sub);
                        bool tempForWeb = web.Exists;
                        bool tempForWeb2 = web.UserIsSiteAdmin;

                        try
                        {
                            //Open library
                            string fullLibraryUrl = web.Url + "/Lists/DocumentLibrary/";
                            SPList list = (SPList)web.GetList(fullLibraryUrl);
                            //SPListItemCollection listItems = list.Items;

                            //get charge accounts data from database
                            string PRQuery =
                            "SELECT " +
                            "[iaa_number] " +
                            "FROM dwmain.extnl.[vLibraryProjectLookup]";

                            DataSet chargeAccountsData = getQueryChargeAccounts(PRQuery);

                            DataTable dataTable = new DataTable();
                            dataTable = chargeAccountsData.Tables[0];

                            SPFieldChoice col1 = (SPFieldChoice)list.Fields["ChargeAccount1"];
                            SPFieldChoice col2 = (SPFieldChoice)list.Fields["ChargeAccount2"];
                            SPFieldChoice col3 = (SPFieldChoice)list.Fields["ChargeAccount3"];
                            SPFieldChoice col4 = (SPFieldChoice)list.Fields["ChargeAccount4"];
                            SPFieldChoice col5 = (SPFieldChoice)list.Fields["ChargeAccount5"];

                            col1.Choices.Clear();
                            col2.Choices.Clear();
                            col3.Choices.Clear();
                            col4.Choices.Clear();
                            col5.Choices.Clear();

                            string firstChoice = ""; 
                            
                            foreach (DataRow dataRow in dataTable.Rows)
                            {
                                if (dataTable.Rows.IndexOf(dataRow) == 0)
                                {
                                    firstChoice = dataRow["iaa_number"].ToString();
                                }
                                col1.Choices.Add(dataRow["iaa_number"].ToString());
                                col2.Choices.Add(dataRow["iaa_number"].ToString());
                                col3.Choices.Add(dataRow["iaa_number"].ToString());
                                col4.Choices.Add(dataRow["iaa_number"].ToString());
                                col5.Choices.Add(dataRow["iaa_number"].ToString());
                            }

                            col1.DefaultValue = firstChoice;
                            col2.DefaultValue = firstChoice;
                            col3.DefaultValue = firstChoice;
                            col4.DefaultValue = firstChoice;
                            col5.DefaultValue = firstChoice; 
                            
                            col1.Update();
                            col2.Update();
                            col3.Update();
                            col4.Update();
                            col5.Update();

                            list.Update();

                        }
                        catch (Exception e)
                        {

                        }
                    }
                }
            });
        }

        public static string StringConnection()
        {
            return "Server=Vdbw001tv\\msapps;Database=product_repository;Trusted_Connection=True;";
        }

        public static DataSet getQuery(string query)
        {
            using (SqlConnection dbConnection = new SqlConnection(StringConnection()))
            {

                dbConnection.Open();
                    
                SqlDataAdapter objCmd = new SqlDataAdapter(query, dbConnection);
                DataSet objDS = new DataSet();
                objCmd.Fill(objDS, "Data");
        
                dbConnection.Close();

                return objDS;        
                    
            }
            
        }

        public static string StringConnectionChargeAccounts()
        {
            return "Server=Vdbw006\\dw;Database=dwmain;Trusted_Connection=True;";
        }

        public static DataSet getQueryChargeAccounts(string query)
        {
            using (SqlConnection dbConnection = new SqlConnection(StringConnectionChargeAccounts()))
            {

                dbConnection.Open();

                SqlDataAdapter objCmd = new SqlDataAdapter(query, dbConnection);
                DataSet objDS = new DataSet();
                objCmd.Fill(objDS, "ChargeAccounts");

                dbConnection.Close();

                return objDS;

            }

        }

    }
}
