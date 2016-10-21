using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;

namespace ConferenceProgramToDocxMapper
{
    public static class JsonHelper
    {
        /// <summary>
        /// download program from conference-publishing.com and parse
        /// </summary>
        /// <param name="uri"></param>
        /// <returns>parsed json object of the download string</returns>
        public static RootObject GetProgramFromWebsite(string uri)
        {
            using (var webClient = new System.Net.WebClient())
            {
                Console.WriteLine("> Fetching json from '{0}'.", uri);
                var jsonString = webClient.DownloadString(uri);
                return JsonConvert.DeserializeObject<RootObject>(jsonString);
            }
        }

        /// <summary>
        /// read from json file and parse
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>parsed json of the file string</returns>
        public static RootObject GetProgramFromFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                var jsonString = File.ReadAllText(filePath, Encoding.UTF8);
                return JsonConvert.DeserializeObject<RootObject>(jsonString);
            }

            throw new Exception("JSON file path is invalid");
        }
    }
}
