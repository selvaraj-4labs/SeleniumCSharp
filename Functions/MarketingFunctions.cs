using OpenQA.Selenium;
using RelevantCodes.ExtentReports;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumCSharpMSTest.ObjectRepository;

namespace SeleniumCSharpMSTest.Functions
{
    public class MarketingFunctions : MarketingObjects
    {

        /// <summary>
        /// Method to Create a new Campaign
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="testInReport"></param>
        /// <param name="testName"></param>
        /// <param name="name"></param>
        /// <param name="campaignCode"></param>
        /// <param name="campaignType"></param>
        /// <param name="expectedResponse"></param>
        /// <param name="proposedStart"></param>
        /// <param name="proposedEnd"></param>
        /// <param name="actualStart"></param>
        /// <param name="actualEnd"></param>
        /// <param name="offer"></param>
        /// <param name="allocatedBudget"></param>
        /// <param name="miscCost"></param>
        /// <param name="description"></param>
        /// <param name="expectedRevenue"></param>
        /// <param name="statusDetails"></param>
        public void CreateNewMarketingCampaign( IWebDriver driver,  ExtentTest testInReport,    string testName,    string testDataIteration,   string name,    string campaigncode,    string campaigntype,    string expectedresponse,    string proposedstart,   string proposedend, string actualstart, string actualend,   string offerData,   string allocatedbudget, string miscCost,    string descriptionData, string expectedrevenue, string statusdetails    )
        {
            try
            {
                // Navigate to Marketing - Campaign
                WaitUntil(driver, tabDropdownMainmenu, 15);
                Element(driver, tabDropdownMainmenu).Click();
                ThinkTime(2);
                Element(driver, marketing).Click();
                ThinkTime(1);
                Element(driver, campaings).Click();
                WaitUntil(driver, campaignNewRecord, 25);
                Element(driver, campaignNewRecord).Click();
                ThinkTime(8);
                SwitchToFrame(driver, "contentIFrame1");
                ThinkTime(2);

                // Enter the Details for the Campaign
                Element(driver, campaignName).Click();
                Element(driver, campaignNameInput).SendKeys(name);
                ThinkTime(1);
                Element(driver, campaignCode).Click();
                Element(driver, campaignCodeInput).SendKeys(campaigncode);
                ThinkTime(1);
                Element(driver, campaignType).Click();
                RunJavaScript(driver, "campaignTypeSelect", campaigntype, "Marketing");
                Element(driver, campaignTypeLabel).Click();

                ThinkTime(1);
                Element(driver, proposedStart).Click();
                ActionSendKeys(driver,  proposedStartInput, proposedstart);
                ThinkTime(1);
                Element(driver, proposedEnd).Click();
                ActionSendKeys(driver,  proposedEndInput, proposedend);
                ThinkTime(1);
                Element(driver, actualStart).Click();
                ActionSendKeys(driver,  actualStartInput, actualstart);
                ThinkTime(1);
                Element(driver, actualEnd).Click();
                ActionSendKeys(driver,  actualEndInput, actualend);
                ThinkTime(1);
                Element(driver, expectedResponse).Click();
                ActionSendKeys(driver,  expectedResponseInput, expectedresponse);
                ThinkTime(1);
                Element(driver, offer).Click();
                ActionSendKeys(driver,  offerInput, offerData);
                ThinkTime(1);
                Element(driver, allocatedBudget).Click();
                ActionSendKeys(driver,  allocatedBudgetInput, allocatedbudget);
                ThinkTime(1);
                Element(driver, miscBudget).Click();
                ActionSendKeys(driver,  miscBudgetInput, miscCost);
                ThinkTime(1);
                Element(driver, description).Click();
                ActionSendKeys(driver,  descriptionsInput, descriptionData);
                ThinkTime(1);
                Element(driver, expectedRevenue).Click();
                ActionSendKeys(driver,  expectedRevenueInput, expectedrevenue);
                ThinkTime(1);
                Element(driver, statusDetails).Click();
                ((IJavaScriptExecutor)driver).ExecuteScript(statusDetailsSelect(statusdetails));
                ThinkTime(1);
                driver.SwitchTo().DefaultContent();
                Element(driver, saveButton).Click();
                ThinkTime(5);

                if (Element(driver, verifyCreatedCampaign(name)).Displayed)
                {
                    ThinkTime(1);
                    AddLog(driver, testInReport, testName, testDataIteration, "Pass", "Campaign successfully created", "Campaign successfully created");
                }
                else
                {
                    ThinkTime(1);
                    AddLog(driver, testInReport, testName, testDataIteration, "Fail", "Campaign creation failed", "Campaign creation failed");
					Assert.IsTrue(false);
                }
            }
            catch (Exception e)
            {
                AddLog(driver, testInReport, testName, testDataIteration, "Fail", "Unexpected error:\n "+e, "Unexpected error");
                throw;
            }
        }
    }
}
