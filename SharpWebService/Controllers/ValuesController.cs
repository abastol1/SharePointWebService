using System;
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


namespace SharpWebService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {

        public string globalFilenameTest = "";

        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }


        //// GET api/values/montvalenames
        //[HttpGet("/filenames/{DocumentLibrary}")]
        //public ActionResult<IEnumerable<string>> Get(string documentLibraryName)
        //{


        //    return new string[] { "value1", "value2" };

        //}

        // GET api/values/5
        [HttpGet("{filename}")]
        public async Task<ActionResult> Get(string filename)
        {
            // making sure any temp documents are not stored on the web api storage
            this.deleteAllTempDocuments();
            globalFilenameTest = filename;

            string searchString = filename.Split('.')[0];
            Debug.WriteLine("Filename received: " + filename);

            byte[] filedata = new byte[0];
            string contentType = "";

            try
                {
                string query = "<mylistitemrequest>" +
                                "<Query>" +
                                    //"<Where>" +
                                    //"<Eq>" +
                                    //    "<FieldRef Name=\"FileLeafRef\" />" +
                                    //    "<Value Type=\"Text\">" + searchString + "</Value>" +
                                    //"</Eq>" +
                                    //"</Where>" +
                                "</Query>" +
                                "<ViewFields>" +
                                    "<FieldRef Name=\"EncodedAbsUrl\"/><FieldRef Name=\"ID\" /><FieldRef Name=\"FileRef\" /><FieldRef Name=\"ID\" /><FieldRef Name=\"Title\" />" +
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

                var retNode = await myService.GetListItemsAsync(
                    Data.documentLibraryName,
                    string.Empty, 
                    queryNode, 
                    viewNode, 
                    string.Empty, 
                    optionNode, 
                    Data.libraryKey
                );

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

                            // ********************************************************************************************
                            filedata = System.IO.File.ReadAllBytes(filePath);
                            contentType = this.GetContentType(filePath);
                            var cd = new System.Net.Mime.ContentDisposition
                            {
                                FileName = globalFilenameTest,
                                Inline = true
                            };
                            Response.Headers["Content-Disposition"] = cd.ToString();
                            // ********************************************************************************************
                        }
                    }
                }
            }
            catch (FaultException fe)
            {
                MessageFault msgFault = fe.CreateMessageFault();
                XmlElement elm = msgFault.GetDetail<XmlElement>();
                
            }
            //return View();
            return new FileStreamResult(new MemoryStream(filedata), contentType);

        }

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
                this.downloadFile(completeFilePath);

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

        private void downloadFile(string filePath)
        {
            try
            {
                //byte[] filedata = System.IO.File.ReadAllBytes(filePath);
                //string contentType = this.GetContentType(filePath);

                //var cd = new System.Net.Mime.ContentDisposition
                //{
                //    FileName = globalFilenameTest,
                //    Inline = true
                //};

                //Response.Headers["Content-Disposition"] = cd.ToString();
                //return new FileStreamResult(new MemoryStream(filedata), contentType);

            }
            catch ( Exception ex )
            {
                Debug.WriteLine(ex.Message);
            }

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
            foreach ( FileInfo file in tempInfo.GetFiles())
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