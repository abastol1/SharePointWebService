﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Xml.XPath;
using System.Xml;
using System.Net;
using System.IO;
using SharpReference;
using System.Xml.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using SharpWebService.Controllers;
using System.Diagnostics;
using System.Configuration;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;
using Newtonsoft.Json;


namespace SharpWebService.Controllers
{
    [Route("")]
    [ApiController]
    public class ValuesController : ControllerBase
    {

        public string globalFilenameTest = "";

        // GET api/values
        [Route("api/[controller]")]
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        /* ***************************************************************************************************************
        Function Name: GetFileNameList
        Purpose: to get list of all the files that are going to be addded on 
        Parameters: none

        Return Value: dictionary with category as key, list of filenames as values
        Local Variables:
                    query, string: CAML Query that specifies which fields to extract from given condition(optional) inside <Query> </Query>
                    listNameDictionary, Dictionary: Stores all the filename categorized by 'Category', key=Category, value=list of filename
        Algorithm:
                    1) Get the details of document libray from SharePoint
                    2) get name of file, category of file and file extension of the file
                    3) Remove .aspx from filename
                    4) if category already exists in the Dictionary, add filename to the value of the category as key
                    5) if categoy doesnot exists in the Dictionary, add category as key, and filename as value
                    6) return the dictionary as string to the calling function
        **************************************************************************************************************** */
        //[Route("fileNameList")]
        [HttpGet("fileNameList")]
        public async Task<string> GetFileNameList()
        {
            string query = "<mylistitemrequest>" +
                "<Query>" + "</Query>" +
                "<ViewFields>" +
                    "<FieldRef Name=\"EncodedAbsUrl\"/><FieldRef Name=\"ID\" /><FieldRef Name=\"MobilePage\" /><FieldRef Name=\"URL\" /><FieldRef Name=\"FileRef\" /><FieldRef Name=\"ID\" /><FieldRef Name=\"Title\" /><FieldRef Name=\"Category\" />" +
                "</ViewFields>" +
                "<QueryOptions></QueryOptions>" +
                "</mylistitemrequest>";

            Dictionary<string, Dictionary<string, List<string>>> mainList = new Dictionary<string, Dictionary<string, List<string>>>();

            Dictionary<string, List<string>> listNameDictionary = new Dictionary<string, List<string>>();

            DataTable dt = null;
            ListsSoapClient.EndpointConfiguration endpoint = new ListsSoapClient.EndpointConfiguration();
            ListsSoapClient myService = new ListsSoapClient(endpoint);
            myService.ClientCredentials.UserName.UserName = Data.EmailUserName;
            myService.ClientCredentials.UserName.Password = Data.EmailCredentials;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(query);

            XElement queryNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//Query")));
            XElement viewNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//ViewFields")));
            XElement optionNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//QueryOptions")));

            var retNode = await myService.GetListItemsAsync(Data.documentLibraryName, string.Empty, queryNode, viewNode, string.Empty, optionNode, Data.libraryKey);

            // Collection of DataTables, stores many datatables in a single collection
            DataSet ds = new DataSet();
            using (StringReader sr = new StringReader(retNode.Body.GetListItemsResult.ToString()))
            {
                ds.ReadXml(sr);
            }

            if (ds.Tables["Row"] != null && ds.Tables["Row"].Rows.Count > 0)
            {
                dt = ds.Tables["Row"].Copy();
                foreach (DataRow dr in dt.Rows)

                {
                    string fileName = dr["ows_FileLeafRef"].ToString().Split("#")[1];
                    string category = dr["ows_Category"].ToString();
                    string page = dr["ows_MobilePage"].ToString();
                    string fileExtension = dr["ows_URL"].ToString().Split(',')[0].Split('.')[1];


                    //if (!page.ToLower().Equals(pageName.ToLower()))
                    //{
                    //    continue;
                    //}

                    if (fileName.Substring(fileName.Length - 5).Equals(".aspx"))
                    {
                        fileName = fileName.Remove(fileName.Length - 5);
                        fileName = fileName + "." + fileExtension;
                    }

                    if (!mainList.ContainsKey(page))
                    {
                        // Create a tempDict, used as value for mainList
                        Dictionary<string, List<string>> tempDict = new Dictionary<string, List<string>>();
                        // Create a tempList, used as value for tempDict
                        List<string> tempList = new List<string>();
                        tempList.Add(fileName);
                        tempDict.Add(category, tempList);
                        mainList.Add(page, tempDict);
                    }
                    else
                    {
                        if (mainList[page].ContainsKey(category))
                        {
                            mainList[page][category].Add(fileName);
                        }
                        else
                        {
                            List<string> tempList = new List<string>();
                            tempList.Add(fileName);
                            mainList[page].Add(category, tempList);
                        }
                    }

                    //if (listNameDictionary.ContainsKey(category))
                    //{
                    //    listNameDictionary[category].Add(fileName);
                    //}
                    //else
                    //{
                    //    List<string> valueFileName = new List<string>();
                    //    valueFileName.Add(fileName);
                    //    listNameDictionary.Add(category, valueFileName);
                    //}
                }
            }
            return JsonConvert.SerializeObject(mainList);
        }


        /* ***********************************************************************************************
        Function Name: Get
        Purpose: to get the file received as input from mobile application 
        Parameters: filename: string, name of the file

        Return Value: returns the file to via HTTP response/ opens file in default browser
        Local Variables:

        Algorithm:
                    1) 
                    2) 
                    3) 
        ************************************************************************************************* */
        // GET api/values/5
        [Route("documents/{filename}")]
        [HttpGet("{filename}")]
        public async Task<ActionResult> Get(string filename)
        {
            // making sure any temp documents are not stored on the web api storage
            this.deleteAllTempDocuments();
            globalFilenameTest = filename;

            byte[] filedata = new byte[0];
            string contentType = "";

            try
            {
                string query = "<mylistitemrequest>" +
                                "<Query>" + "</Query>" +
                                "<ViewFields>" +
                                    "<FieldRef Name=\"EncodedAbsUrl\"/><FieldRef Name=\"ID\" /><FieldRef Name=\"FileRef\" /><FieldRef Name=\"ID\" /><FieldRef Name=\"Title\" /><FieldRef Name=\"Category\" />" +
                                "</ViewFields>" +
                                "<QueryOptions></QueryOptions>" +
                                "</mylistitemrequest>";
                DataTable dt = null;
                ListsSoapClient.EndpointConfiguration endpoint = new ListsSoapClient.EndpointConfiguration();

                ListsSoapClient myService = new ListsSoapClient(endpoint);
                myService.ClientCredentials.UserName.UserName = Data.EmailUserName;
                myService.ClientCredentials.UserName.Password = Data.EmailCredentials;
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(query);

                XElement queryNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//Query")));
                XElement viewNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//ViewFields")));
                XElement optionNode = XElement.Load(new XmlNodeReader(doc.SelectSingleNode("//QueryOptions")));

                // Data.documentLibraryName: Display name of the document library
                // string.Empty: ViewName
                // queryNode: query element containing the query that determines which records are returned and in what order
                // viewNode: element that specifies which fields to return in the query and in what order
                // optionNode: An XML fragment that contains separate nodes for the various properties
                // Data.libraryKey: string containing the GUID of the parent Website for the list
                // Return Value: Returns information about items in the list based on the specified query
                var retNode = await myService.GetListItemsAsync(Data.documentLibraryName, string.Empty, queryNode, viewNode, string.Empty, optionNode, Data.libraryKey);

                // Collection of DataTables, stores many datatables in a single collection
                DataSet ds = new DataSet();
                using (StringReader sr = new StringReader(retNode.Body.GetListItemsResult.ToString()))
                {
                    ds.ReadXml(sr);
                }

                if (ds.Tables["Row"] != null && ds.Tables["Row"].Rows.Count > 0)
                {
                    dt = ds.Tables["Row"].Copy();
                    foreach (DataRow dr in dt.Rows)
                    {
                        string siteURL = dr["ows_EncodedAbsUrl"].ToString();
                        string fileName = dr["ows_FileLeafRef"].ToString().Split("#")[1];

                        if (fileName.Contains(".aspx"))
                        {
                            fileName = fileName.Replace(".aspx", ".pdf");
                        }
                        if (globalFilenameTest.Equals(fileName))
                        {
                            string filePath = DownLoadAttachment(dr["ows_EncodedAbsUrl"].ToString(), fileName);
                            filedata = System.IO.File.ReadAllBytes(filePath);
                            contentType = this.GetContentType(filePath);
                            var cd = new System.Net.Mime.ContentDisposition
                            {
                                FileName = globalFilenameTest,
                                Inline = true
                            };
                            Response.Headers["Content-Disposition"] = cd.ToString();
                        }
                    }
                }
            }
            catch (FaultException fe)
            {
                MessageFault msgFault = fe.CreateMessageFault();
                XmlElement elm = msgFault.GetDetail<XmlElement>();
            }
            return new FileStreamResult(new MemoryStream(filedata), contentType);
        }


        /* *************************************************************************************************
        Function Name: DownloadAttachment
        Purpose: downloads file temporarily in local storage from SharePoint Document Library
        Parameters: strURL, string: 

        Return Value: returns the filepath
        Local Variables:
                completeFilePath, string: Filepath where the temp document is stored
                workingDirectory, string: Location where the temp document is created
        Algorithm:
                1) Get ResponseStream from Sharepoint Document URL
                2) Read the content using the url and store it in a byte array
                3) Create a file in the "workingDirectory" location
                4) write content to the newly created Document
                5) return the path where the file was created
        ***************************************************************************************************** */
        public string DownLoadAttachment(string strURL, string strFileName)
        {
            string completeFilePath = "";
            HttpWebRequest request;
            HttpWebResponse response = null;
            try
            {
                request = (HttpWebRequest)WebRequest.Create(strURL);
                request.Credentials = System.Net.CredentialCache.DefaultCredentials;
                request.Timeout = 10000;
                request.AllowWriteStreamBuffering = false;
                response = (HttpWebResponse)request.GetResponse();
                Stream s = response.GetResponseStream();
                //Write to disk
                string workingDirectory = Directory.GetCurrentDirectory().ToString() + @"\wwwroot\tempDocuments\";
                completeFilePath = workingDirectory + strFileName;
                FileStream fs = new FileStream(workingDirectory + strFileName, FileMode.Create);
                byte[] read = new byte[256];
                int count = s.Read(read, 0, read.Length);
                while (count > 0)
                {
                    fs.Write(read, 0, count);
                    count = s.Read(read, 0, read.Length);
                }

                //Close everything
                fs.Close();
                s.Close();
                response.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return completeFilePath;
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        /* *********************************************************************
        Function Name: GetContentType
        Purpose: to get the content type of a file
        Parameters: path: string, an absolute path of a file

        Return Value: content type
        Local Variables:
                    types: Dictionary<Key, value> of mimetypes
                    ext: extension of a file
        Algorithm:
                    1) get mimetypes from GetMimeTypes() function
                    2) get only the extension of file
                    3) using ext as key, return the value
        ********************************************************************* */
        private string GetContentType(string path)
        {
            var types = GetMimeTypes();
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return types[ext];
        }


        /* *********************************************************************
        Function Name: GetMimeTypes
        Purpose: to get the dictionary of mime types where key=file extension, value=mimetype
        Parameters:

        Return Value: Collection of key/value pairs of mimetypes
        Local Variables:
                    none
        Algorithm:
                    1) return the dictionary
        ********************************************************************* */
        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                { ".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.openxmlformatsofficedocument.spreadsheetml.sheet"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
                {".jpeg", "image/jpeg"},
                {".gif", "image/gif"},
                {".csv", "text/csv"}
            };
        }

        /* *********************************************************************
        Function Name: deleteAllTempDocuments
        Purpose: To delete all the documents in TempDocuments directory that might be download when accessing SharePoint API
        Parameters: None

        Return Value: none
        Local Variables:
                    tempInfo: DirectoryInfo, Get the info of tempDocuments directory, also contains name of files
        Algorithm:
                    1) Get info of tempDocuments
                    2) Loop throught all the files in the directory
                        3) Delete the file
        ********************************************************************* */
        private void deleteAllTempDocuments()
        {
            DirectoryInfo tempInfo = new DirectoryInfo(Directory.GetCurrentDirectory().ToString() + @"\wwwroot\tempDocuments\");
            foreach (FileInfo file in tempInfo.GetFiles())
            {
                file.Delete();
            }
        }
    }
}

//"<Eq>" +
//    "<FieldRef Name=\"IsOnApp\" />" +
//    "<Value Type=\"Lookup\">" + "1" + "</Value>" +
//"</Eq>" +

//"<Where>" +
//"<Eq>" +
//    "<FieldRef Name=\"FileLeafRef\" />" +
//    "<Value Type=\"Text\">" + searchString + "</Value>" +
//"</Eq>" +
//"</Where>" +