
<%@ WebHandler Language="C#" Class="SPFormFileHandler" %>

using System;
using System.Web;
using System.IO;
using System.Text;
using Microsoft.SharePoint;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;


//temporarily elevate directory write privileges
public class SPFormFileHandler : IHttpHandler
{
    public void ProcessRequest(HttpContext context)
    {
        context.Response.CacheControl = "no-cache";
        context.Response.AddHeader("Pragma", "no-cache");
        context.Response.Expires = -1;
        context.Response.ContentType = "text/plain";
        context.Response.ContentEncoding = UTF8Encoding.UTF8;

        HttpRequest r = context.Request;
        Response res = new Response(false, null, null, null);
        Guid listId = new Guid();
        int? itemId = null;
        string fileName = null;
        byte[] fileContent;
        StringBuilder errors = new StringBuilder();

        try
        {
            if (!string.IsNullOrEmpty(r["listId"]))
                listId = new Guid(r["listId"].ToString());
            else
                errors.Append("Missing List ID. ");

            if (!string.IsNullOrEmpty(r["itemId"]) && r["itemId"] != "null")
                itemId = Convert.ToInt32(r["itemId"].ToString());

            //not IE
            if (!string.IsNullOrEmpty(r["qqfile"]))
            {
                fileName = r["qqfile"].ToString();
                fileName = CleanFileName(Path.GetFileName(fileName));
                fileContent = StreamToByteArray(r.InputStream);
                res = Import(listId, itemId, fileName, fileContent);
            }
            else //is IE
            {
                HttpFileCollection hfc = r.Files;
                for (int i = 0; i < hfc.Count; i++)
                {
                    HttpPostedFile hpf = hfc[i];
                    if (hpf.ContentLength > 0)
                    {
                        fileName = CleanFileName(Path.GetFileName(hpf.FileName));
                        fileContent = StreamToByteArray(hpf.InputStream);
                        res = Import(listId, itemId, fileName, fileContent);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            errors.AppendFormat("Item ID: {0}, List ID: {1}, Filename: {2}, Error: {3}", itemId, listId.ToString(), fileName, ex.ToString());
        }
        finally
        {
            res.error = errors.ToString();
            string json = new JavaScriptSerializer().Serialize(res);
            context.Response.Write(json);
        }   
    }

    public static Response Import(Guid listId, int? itemId, string fileName, byte[] fileContent)
    {
        Response res = new Response(false, null, null, null);
        try
        {
            using (SPWeb webContext = SPContext.Current.Web)
            {
                using (SPSite site = new SPSite(webContext.Site.ID))
                {
                    using (SPWeb web = site.OpenWeb(webContext.ID))
                    {
                        web.AllowUnsafeUpdates = true;
                        try
                        {
                            SPList list = web.Lists[listId];
                            SPListItem item = itemId != null ? list.GetItemById((int)itemId) : list.AddItem();

                            item.Attachments.Add(fileName, fileContent);
                            item.Update();
                            res.success = true;
                            res.fileName = fileName;
                            res.itemId = item.ID;
                        }
                        catch (Exception ex)
                        {
                            res.success = false;
                            res.error = ex.Message;
                        }
                        finally
                        {
                            web.AllowUnsafeUpdates = false;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            res.success = false;
            res.error = ex.ToString();
        }

        return res;
    }

    public class Response
    {
        private bool _success;
        private string _error;
        private int? _itemId;
        private string _fileName;
        public bool success { get { return this._success; } set { this._success = value; } }
        public string error { get { return this._error; } set { this._error = value; } }
        public int? itemId { get { return this._itemId; } set { this._itemId = value; } }
        public string fileName { get { return this._fileName; } set { this._fileName = value; } }

        public Response(bool _success, string _error, int? _itemId, string _fileName)
        {
            this._success = _success;
            this._error = _error;
            this._itemId = _itemId;
            this._fileName = _fileName;
        }
    }

    public static string CleanFileName(string filename)
    {
        filename = Regex.Replace(filename, "[^a-zA-Z0-9._]", "");
        filename = Regex.Replace(filename, "[\\^{}]", "");
        filename = filename.Replace("[", "").Replace("]", "");
        return filename;
    }

    private static byte[] StreamToByteArray(Stream inputStream)
    {

        if (inputStream.CanSeek)
            inputStream.Seek(0, SeekOrigin.Begin);

        byte[] output = new byte[inputStream.Length];
        int bytesRead = inputStream.Read(output, 0, output.Length);
        return output;
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }
}
