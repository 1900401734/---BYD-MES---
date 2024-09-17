using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace MesDatas
{
    public class XmlSerializerHelper
    {
        /// <summary>
        /// 反序列化xml文件，并检查加密是否正确
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strPath"></param>
        /// <param name="strErrMessage"></param>
        /// <param name="IsDecode"></param>
        /// <returns></returns>
        public static T Load<T>(string strPath, ref string strErrMessage)
        {
            strErrMessage = string.Empty;
            T config = default(T);
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(T));
                FileStream fs = new FileStream(strPath, FileMode.Open, FileAccess.Read);
                config = (T)xs.Deserialize(fs);
                fs.Close();
            }
            catch (Exception err)
            {
                Debug.WriteLine(err.ToString());
                strErrMessage += err.ToString();
            }
            return config;
        }

        public static void SaveUtf8NoNamespace<T>(string strFile, T obj, ref string strErrMessage)
        {
            strErrMessage = string.Empty;

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.Indent = true;
            try
            {
                using (XmlWriter writer = XmlWriter.Create(strFile, settings))
                {
                    XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                    ns.Add("", "");
                    XmlSerializer formatter = new XmlSerializer(typeof(T));
                    formatter.Serialize(writer, obj, ns);
                    writer.Close();
                }
            }
            catch (Exception err)
            {
                Debug.WriteLine(err.ToString());
                strErrMessage += err.ToString();
            }
        }

        public static void Save<T>(string strPath, T obj, ref string strErrMessage)
        {
            if (obj != null)
            {
                FileStream fs = null;
                try
                {
                    XmlSerializer xs = new XmlSerializer(typeof(T));
                    fs = new FileStream(strPath, FileMode.Create, FileAccess.Write);
                    xs.Serialize(fs, obj);
                    fs.Close();
                }
                catch (Exception err)
                {
                    Debug.WriteLine(err.ToString());
                    fs.Close();
                    strErrMessage += err.ToString();
                }
            }
        }

        public static bool Save(object objToSave, string strFilePath)
        {
            bool bRet = false;
            FileStream s = null;
            try
            {
                s = File.Open(strFilePath, FileMode.Create, FileAccess.ReadWrite);
                XmlSerializer sr = new XmlSerializer(objToSave.GetType());
                sr.Serialize(s, objToSave);
                bRet = true;
            }
            catch (Exception err)
            {
                throw err;
            }
            finally
            {
                if (null != s)
                {
                    s.Close();
                }
            }
            return bRet;
        }

        public static T DeserializeXML<T>(string xmlObj)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (StringReader reader = new StringReader(xmlObj))
            {
                return (T)serializer.Deserialize(reader);
            }
        }

        public static string SerializeXML<T>(T obj)
        {
            using (StringWriter writer = new StringWriter())
            {
                new XmlSerializer(obj.GetType()).Serialize(writer, obj);
                return writer.ToString();
            }
        }

        public static void AddXmlComment(string strFile, string strCommet)
        {
            if (File.Exists(strFile))
            {
                using (StreamWriter sw = new StreamWriter(strFile, true, Encoding.Default))
                {
                    sw.WriteLine("");
                    sw.WriteLine(CommentFront + strCommet + CommentBack);
                }
            }
        }

        public static List<string> GetCommetFromXml(string strFile)
        {
            List<string> lsRet = new List<string>();
            if (File.Exists(strFile))
            {
                XmlReader xmlReader = XmlReader.Create(strFile);
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Comment)
                    {
                        lsRet.Add(xmlReader.Value);
                    }
                }
            }
            return lsRet;
        }

        public static string Serialize<T>(T obj)
        {
            XmlSerializer ser = Create<T>();
            MemoryStream stream = new MemoryStream();
            ser.Serialize(stream, obj);
            byte[] buffer = stream.ToArray();
            return Encoding.UTF8.GetString(buffer);
        }


        public static T DeserializeString<T>(string xml)
        {
            XmlSerializer ser = Create<T>();
            byte[] buffer = Encoding.UTF8.GetBytes(xml);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(buffer, 0, buffer.Length);
                stream.Position = 0L;
                using (XmlReader reader = XmlReader.Create(stream))
                {
                    return (T)ser.Deserialize(reader);
                }
            }
        }

        private static XmlSerializer Create<T>()
        {
            return new XmlSerializer(typeof(T));
        }

        public static string ReadXml(string file)
        {
            using (FileStream stream = new FileStream(file, FileMode.Open))
            {
                StreamReader reader = new StreamReader(stream);
                string content = reader.ReadToEnd();
                stream.Close();
                return content;
            }
        }


        private const string CommentFront = "<!--";
        private const string CommentBack = "-->";

    }
}
