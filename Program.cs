﻿using System;
using HtmlAgilityPack;
using ScrapySharp.Extensions;
using CsvHelper;
using System.IO;
using System.Net;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Threading;

namespace web_scrap
{
    
        class Program
    {
        static void Main(string[] args)
        {
            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            var options = getCurrencyOptions();
            options = options.Where(o => o.Length == 3);


            foreach(var opt in options){
               Thread thread = new Thread(() => scrapeCurrency(opt));
               thread.Start();
            }
            
        }
        static string goToPage(String currencyInput,int count,string first_date, string curr_date){
            try{
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://srh.bankofchina.com/search/whpj/searchen.jsp");
            request.Method = WebRequestMethods.Http.Post;
            request.ContentType = "application/x-www-form-urlencoded";

           
            var encoding=new ASCIIEncoding();
            var postData="erectDate="+first_date;
            postData += ("&nothing="+curr_date);
            postData += ("&pjname="+currencyInput);
            postData += ("&page="+count);
            byte[]  data = encoding.GetBytes(postData);
            request.ContentLength = data.Length;

            var newStream=request.GetRequestStream();
            newStream.Write(data,0,data.Length);
            
            HttpWebResponse myResponse = (HttpWebResponse)request.GetResponse();
            string result = "";
            StreamReader reader = new StreamReader(myResponse.GetResponseStream());
            result = reader.ReadToEnd();
            newStream.Close();
            return result;
             }
            catch(Exception ex){
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        public static IEnumerable<string> getCurrencyOptions(){

            HtmlWeb web = new HtmlWeb();
            var htmlDoc = web.Load("https://srh.bankofchina.com/search/whpj/searchen.jsp");
            var options = htmlDoc.GetElementbyId("pjname").CssSelect("option").Select(l => l.InnerText);
            return options;

        }
        public static void scrapeCurrency(string currencyInput){


            DateTime dateTime = DateTime.UtcNow.Date;
            var curr_date =  dateTime.ToString("yyyy-MM-dd");
            var first_date = dateTime.AddDays(-2).ToString("yyyy-MM-dd");
            string result = goToPage(currencyInput,1,first_date,curr_date);
            
            try{
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(result);       
            var html = doc.DocumentNode;

            using (var sr = new StreamWriter(currencyInput+"-"+first_date+"-"+curr_date+".csv")) {
            using (var csv = new CsvWriter(sr,System.Globalization.CultureInfo.CurrentCulture)) {
                csv.WriteField("CurrencyName");
                csv.WriteField("BuyingRate");
                csv.WriteField("CashBuyingRate");
                csv.WriteField("SellingRate");
                csv.WriteField("CashSellingRate");
                csv.WriteField("MiddleRate");
                csv.WriteField("PubTime");
                csv.NextRecord();
                int count = 2;
                var pom = result;
                while(true){
                    var trs = html.CssSelect("tr");
                    if(trs.Count()<7){
                        break;
                    }
                    foreach (var tr in trs)
                    {

                        var tds = tr.CssSelect("td");
                        
                        foreach(var td in tds){

                            if (td.GetAttributeValue("class").Equals("hui12_20"))
                            {
                                csv.WriteField(td.InnerText);
                            }
                            
                        }
                        csv.NextRecord();
                        sr.Flush();               
                    }

                   result = goToPage(currencyInput,count,first_date,curr_date);
                    
                    
                    if(result.Equals(pom)){
                       break;
                    }

                    Console.WriteLine("count of scraped pages: " +count + " of: " + currencyInput );
                    pom = result;
                    count++;
                }
              }
            }

            }catch(Exception ex){
                Console.WriteLine(ex.Message);
            }
        }
  
    }
     
}
