using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace WPFtoPTTLibrary
{
    public class ImageProcessor
    {
        //returns ImageModels to be used for images scources for slides
        public static async Task<ImagesModel> LoadImages(string searchTerms)
        {
            string url = $"https://www.googleapis.com/customsearch/v1?q={ searchTerms }&num=8&cx={ConfigurationManager.AppSettings["EngineID"]}&searchType=image&key={ConfigurationManager.AppSettings["ApiKey"]}&alt=json";

            using (HttpResponseMessage response = await ApiHelper.ApiClient.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    ImagesModel images = await response.Content.ReadAsAsync<ImagesModel>();

                    return images;
                }
                else
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }
        }
    }
}
