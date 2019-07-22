using System;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace SlideGen.Services
{
    /// <summary>
    /// The ImageService interfaces with the Unsplash image API to retrieve image uri locations.
    /// The only data that is needed is the url for the small images sizes. 
    /// </summary>
    class ImageService
    {
        private string baseURL = "https://api.unsplash.com/search/photos";
        private string client_id = "40ae8488e79841686ad57cc16c165139c58d76d64c57f6eacea28b652e6bc3ea";

        public ImageService() {}

        /// <summary>
        /// Makes an HTTP GET request to the Unsplash photo search API, and parses the response as a JSON object.
        /// </summary>
        /// <param name="query"></param>
        /// <returns>Array of parsed image URLs</returns>
        public string[] GetImages(string query)
        {
            using (var httpClient = new HttpClient())
            {
                var response = httpClient.GetStringAsync(new Uri(baseURL + "?page=1&query=" + query + "&client_id=" + client_id)).Result;
                var dataObj = JObject.Parse(response);
                return parseImgURIs(dataObj["results"]);
            }
        }

        /// <summary>
        /// Parses the first 5 image results.
        /// </summary>
        /// <param name="results"></param>
        /// <returns>Array of small image size URLs</returns>
        private string[] parseImgURIs(JToken results)
        {
            string[] imgList = new string[Constants.maxImages];

            int i = 0;
            foreach (var item in results)
            {
                imgList[i] = $"{item["urls"]["small"]}";
                ++i;
                if(i >= Constants.maxImages)
                {
                    break;
                }
            }

            return imgList;
        }

    }
}
