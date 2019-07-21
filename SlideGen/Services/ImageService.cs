using System;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace SlideGen.Services
{
    class ImageService
    {
        private const int maxImages = 5; 
        private string baseURL = "https://api.unsplash.com/search/photos";
        private string client_id = "40ae8488e79841686ad57cc16c165139c58d76d64c57f6eacea28b652e6bc3ea";

        public ImageService() {}

        public string[] GetImages(string query)
        {
            using (var httpClient = new HttpClient())
            {
                //httpClient.DefaultRequestHeaders.Add(RequestConstants.UserAgent, RequestConstants.UserAgentValue);

                var response = httpClient.GetStringAsync(new Uri(baseURL + "?page=1&query=" + query + "&client_id=" + client_id)).Result;

                var dataObj = JObject.Parse(response);

                return parseImgURIs(dataObj["results"]);
            }
        }

        // Parse the first 5 image results and return an array of the thumbnail URIs
        private string[] parseImgURIs(JToken results)
        {
            string[] imgList = new string[maxImages];

            int i = 0;
            foreach (var item in results)
            {
                imgList[i] = $"{item["urls"]["small"]}";
                ++i;
                if(i >= maxImages)
                {
                    break;
                }
            }

            return imgList;
        }

    }
}
