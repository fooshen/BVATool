using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.IO.Compression;
using System.Collections.Specialized;
using System.Net.Mail;
using System.Text;
using Microsoft.AspNetCore.Mvc.Filters;
using System.Net;
using System.Collections.Generic;
using System.Threading;

namespace BVADocGen
{
    public static class ProcessWordDoc
    {
        [FunctionName("ProcessWordDoc")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {            
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            string file = data?.file;

            string data_reducedDevCost = data?.data_reducedDevCost;
            string data_ongoingCost = data?.data_ongoingCost;
            string data_efficiencyCost = data?.data_efficiencyCost;
            string data_3partyCost = data?.data_3partyCost;
            string data_runCost = data?.data_runCost;
                       
            ActionResult result;
            if (file != null)
            {
                
                Byte[] content = Convert.FromBase64String(file);
                MemoryStream inputDocx = new MemoryStream();
                inputDocx.Write(content, 0, content.Length);

                string _chart1Categories = @"<c:cat><c:numRef><c:f>'3. Summary chart'!$B$5:$B$10</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""6""/><c:pt idx=""0""><c:v>0</c:v></c:pt><c:pt idx=""1""><c:v>0</c:v></c:pt><c:pt idx=""2""><c:v>0</c:v></c:pt><c:pt idx=""3""><c:v>0</c:v></c:pt><c:pt idx=""4""><c:v>0</c:v></c:pt><c:pt idx=""5""><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:cat>";
                string _chart1Values = @"<c:val><c:numRef><c:f>'3. Summary chart'!$C$5:$C$10</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""6""/><c:pt idx=""0""><c:v>0</c:v></c:pt><c:pt idx=""1""><c:v>0</c:v></c:pt><c:pt idx=""2""><c:v>0</c:v></c:pt><c:pt idx=""3""><c:v>0</c:v></c:pt><c:pt idx=""4""><c:v>0</c:v></c:pt><c:pt idx=""5""><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>";

                string _chart2Categories = @"<c:cat><c:numRef><c:f>'[1]3'!$B$5:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""7""/><c:pt idx=""0""><c:v>0</c:v></c:pt><c:pt idx=""1""><c:v>0</c:v></c:pt><c:pt idx=""2""><c:v>0</c:v></c:pt><c:pt idx=""3""><c:v>0</c:v></c:pt><c:pt idx=""4""><c:v>0</c:v></c:pt><c:pt idx=""5""><c:v>0</c:v></c:pt><c:pt idx=""6""><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:cat>";
                //string _chart2Values = @"<c:val><c:numRef><c:f>'[1]3'!$B$5:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""6""/><c:pt idx=""0""><c:v>0</c:v></c:pt><c:pt idx=""1""><c:v>0</c:v></c:pt><c:pt idx=""2""><c:v>0</c:v></c:pt><c:pt idx=""3""><c:v>0</c:v></c:pt><c:pt idx=""4""><c:v>0</c:v></c:pt><c:pt idx=""5""><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>";
                string _chart2Values = @"<c:val><c:numRef><c:f>'[1]3'!$C$5:$C$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""7""/><c:pt idx=""0""><c:v>0</c:v></c:pt><c:pt idx=""1""><c:v>0</c:v></c:pt><c:pt idx=""2""><c:v>0</c:v></c:pt><c:pt idx=""3""><c:v>0</c:v></c:pt><c:pt idx=""4""><c:v>0</c:v></c:pt><c:pt idx=""5""><c:v>0</c:v></c:pt><c:pt idx=""6""><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>";
                try
                {
                    using (ZipArchive zipFile = new ZipArchive(inputDocx, ZipArchiveMode.Update))
                    {                        
                        ZipArchiveEntry chart1Xml = zipFile.GetEntry(@"word/charts/chart1.xml");
                        ZipArchiveEntry chart2Xml = zipFile.GetEntry(@"word/charts/chart2.xml");

                        StringBuilder chart1XmlContent;
                        StringBuilder chart2XmlContent;
                        
                        using (StreamReader chart1Reader = new StreamReader(chart1Xml.Open()))
                        {
                            chart1XmlContent = new StringBuilder(chart1Reader.ReadToEnd());
                        }
                        
                        chart1Xml.Delete();                       
                        chart1Xml = zipFile.CreateEntry(@"word/charts/chart1.xml");                       
                        chart1XmlContent.Replace(_chart1Categories,
                            @"<c:cat><c:strRef><c:f>'3. Summary chart'!$B$5:$B$10</c:f><c:strCache><c:ptCount val=""6""/><c:pt idx=""0""><c:v>Reduced Development Costs</c:v></c:pt><c:pt idx=""1""><c:v>Reduced Ongoing Support</c:v></c:pt><c:pt idx=""2""><c:v>Efficiency Cost Savings</c:v></c:pt><c:pt idx=""3""><c:v>3rd Party App Costs Avoided</c:v></c:pt><c:pt idx=""4""><c:v>Run-Costs Avoided</c:v></c:pt></c:strCache></c:strRef></c:cat>");
                        chart1XmlContent.Replace(_chart1Values,
                            $@"<c:val><c:numRef><c:f>'3. Summary chart'!$C$5:$C$10</c:f><c:numCache><c:formatCode>""$""#,##0,""K""</c:formatCode><c:ptCount val=""6""/><c:pt idx=""0""><c:v>{ data_reducedDevCost }</c:v></c:pt><c:pt idx=""1""><c:v>{ data_ongoingCost }</c:v></c:pt><c:pt idx=""2""><c:v>{ data_efficiencyCost }</c:v></c:pt><c:pt idx=""3""><c:v>{ data_3partyCost }</c:v></c:pt><c:pt idx=""4""><c:v>{ data_runCost }</c:v></c:pt></c:numCache></c:numRef></c:val>");

                        using (StreamWriter chart1Writer = new StreamWriter(chart1Xml.Open()))
                        {
                            chart1Writer.Write(chart1XmlContent);
                        }
                        
                        using (StreamReader chart2Reader = new StreamReader(chart2Xml.Open()))
                        {
                            chart2XmlContent = new StringBuilder(chart2Reader.ReadToEnd());
                        }
                       
                        chart2Xml.Delete();
                        chart2Xml = zipFile.CreateEntry(@"word/charts/chart2.xml");
                        chart2XmlContent.Replace(_chart2Categories,
                           @"<c:cat><c:strRef><c:f>'[1]3'!$B$5:$B$11</c:f><c:strCache><c:ptCount val=""6""/><c:pt idx=""0""><c:v>Reduced Development Costs</c:v></c:pt><c:pt idx=""1""><c:v>Reduced Ongoing Support</c:v></c:pt><c:pt idx=""2""><c:v>Efficiency Cost Savings</c:v></c:pt><c:pt idx=""3""><c:v>3rd Party App Costs Avoided</c:v></c:pt><c:pt idx=""4""><c:v>Run-Costs Avoided</c:v></c:pt><c:pt idx=""5""><c:v>Total</c:v></c:pt></c:strCache></c:strRef></c:cat>");
                        chart2XmlContent.Replace(_chart2Values,
                            $@"<c:val><c:numRef><c:f>'[1]3'!$C$5:$C$11</c:f><c:numCache><c:formatCode>""$""#,##0,""K""</c:formatCode><c:ptCount val=""6""/><c:pt idx=""0""><c:v>{ data_reducedDevCost }</c:v></c:pt><c:pt idx=""1""><c:v>{ data_ongoingCost }</c:v></c:pt><c:pt idx=""2""><c:v>{ data_efficiencyCost }</c:v></c:pt><c:pt idx=""3""><c:v>{ data_3partyCost }</c:v></c:pt><c:pt idx=""4""><c:v>{ data_runCost }</c:v></c:pt><c:pt idx=""5""><c:v>{ float.Parse(data_reducedDevCost) + float.Parse(data_ongoingCost) + float.Parse(data_efficiencyCost) + float.Parse(data_3partyCost) + float.Parse(data_runCost) }</c:v></c:pt></c:numCache></c:numRef></c:val>");

                        using (StreamWriter chart2Writer = new StreamWriter(chart2Xml.Open()))
                        {
                            chart2Writer.Write(chart2XmlContent);
                        }
                    }
                }
                catch(Exception ex)
                {
                    string message = ex.Message;
                }                
                
                result = (ActionResult)new FileContentResult(inputDocx.ToArray(), @"application/zip");
                //result = (ActionResult)new FileContentResult(content, @"application/zip");
            }
            else
            {
                result = (ActionResult)new BadRequestObjectResult("File is invalid");
            }
            return result;
        }        


    }

    public static class ProcessPowerPoint
    {
        [FunctionName("ProcessPowerPoint")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            string file = data?.file;

            Dictionary<string, string> chartItems = new Dictionary<string, string>();
            chartItems.Add("1001", (string)data.data_reducedDevCost);
            chartItems.Add("1002", (string)data.data_ongoingCost);
            chartItems.Add("1003", (string)data.data_efficiencyCost);
            chartItems.Add("1004", (string)data.data_3partyCost);
            chartItems.Add("1005", (string)data.data_runCost);
            float total = float.Parse((string)data.data_reducedDevCost) + float.Parse((string)data.data_ongoingCost) + float.Parse((string)data.data_efficiencyCost) +
                float.Parse((string)data.data_3partyCost) + float.Parse((string)data.data_runCost);
            chartItems.Add("1006", total.ToString());

            Dictionary<string, string> tokens = new Dictionary<string, string>();
            tokens.Add("_CUSTOMER_", (string) data.customerName);
            tokens.Add("_Presenter_", (string) data.presenterName);
            tokens.Add("_Email_", (string)data.email);
            tokens.Add("_ab_", (string)data.annualAvgBenefit);
            tokens.Add("_roi_", (string)data.roi);
            tokens.Add("_nb_", (string)data.netBenefit);
            tokens.Add("_pyb_", (string)data.payBack);
            tokens.Add("_usr_", (string)data.numUsers);
            tokens.Add("_pa_", (string)data.numPA);
            tokens.Add("_dmc_", (string)data.devCostMedComp);
            tokens.Add("_mc_", (string)data.totalMedCompApps);
            tokens.Add("_bp_", (string)data.bizProcessApps);
            tokens.Add("_fte_", (string)data.numFTE);
            tokens.Add("_licplan_", (string)data.licenseType);
            tokens.Add("_lc_", (string)data.licensingCost);
            tokens.Add("_tdmc_", (string)data.totalDevCostMedComp);
            tokens.Add("_tdbp_", (string)data.totalDevCostProcImp);
            tokens.Add("_ogs_", (string)data.ongoingSuport);
            tokens.Add("_ootc_", (string)data.otherOneTimeCosts);
            tokens.Add("_dbp_", (string)data.devCostBizProc);

            ActionResult result;
            if (file != null)
            {
                Byte[] content = Convert.FromBase64String(file);
                MemoryStream inputPptx = new MemoryStream();
                inputPptx.Write(content, 0, content.Length);

                try
                {
                    using (ZipArchive zipFile = new ZipArchive(inputPptx, ZipArchiveMode.Update))
                    {
                        //process only from slide 1 to slide 6
                        for(int i = 1; i <= 6; i++)
                        {
                            string xmlPath = string.Format(@"ppt/slides/slide{0}.xml", i);
                            ZipArchiveEntry slideXml = zipFile.GetEntry(xmlPath);
                            StringBuilder slideContent;

                            using (StreamReader slideReader = new StreamReader(slideXml.Open()))
                            {
                                slideContent = new StringBuilder(slideReader.ReadToEnd());
                            }

                            slideXml.Delete();
                            slideXml = zipFile.CreateEntry(xmlPath);
                            foreach(string key in tokens.Keys) slideContent.Replace(key, tokens[key]);
                         
                            using (StreamWriter slideWriter = new StreamWriter(slideXml.Open()))
                            {
                                slideWriter.Write(slideContent);
                            }
                        }

                        //update chart items
                        for (int i = 1; i <= 2; i++)
                        {
                            string chartXmlPath = string.Format(@"ppt/charts/chart{0}.xml", i);
                            ZipArchiveEntry chartXml = zipFile.GetEntry(chartXmlPath);
                            StringBuilder chartXmlContent;
                            using (StreamReader chartReader = new StreamReader(chartXml.Open()))
                            {
                                chartXmlContent = new StringBuilder(chartReader.ReadToEnd());
                            }

                            chartXml.Delete();
                            chartXml = zipFile.CreateEntry(chartXmlPath);
                            foreach (string key in chartItems.Keys) chartXmlContent.Replace(key, chartItems[key]);

                            using (StreamWriter chartWriter = new StreamWriter(chartXml.Open()))
                            {
                                chartWriter.Write(chartXmlContent);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string message = ex.Message;
                }

                result = (ActionResult)new FileContentResult(inputPptx.ToArray(), @"application/zip");
                //result = (ActionResult)new FileContentResult(content, @"application/zip");
            }
            else
            {
                result = (ActionResult)new BadRequestObjectResult("File is invalid");
            }
            return result;
        }


    }
}
