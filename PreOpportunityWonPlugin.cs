using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;


namespace ASG.CRM.PreOpportunityWonPlugin
{
    public class PreOpportunityWonPlugin : IPlugin
    {
        public IOrganizationService wService;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region variables

            Guid guidOpportunityId = new Guid();
            List<Quote> quoteToBeClosed = new List<Quote>();

            #endregion

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));


            if (context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity)
            {
                // Get a refrence to CRM API Services
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                wService = serviceFactory.CreateOrganizationService(context.UserId);

                //create entity context
                Entity entity = (Entity)context.InputParameters["OpportunityClose"];

                if (entity.LogicalName != OpportunityClose.EntityLogicalName)
                {
                    return;
                }
                try
                {
                    //create target entity as early bound
                    OpportunityClose TargetEntity = entity.ToEntity<OpportunityClose>();
 
                    //get the Opportunity Id
                    guidOpportunityId = TargetEntity.OpportunityId.Id;

                    //get all quote related to this Opportunity
                    quoteToBeClosed = GetQuotes(guidOpportunityId);

                    // Make sure there is a won quote
                    int iWonQuotes = 0;
                    if (quoteToBeClosed != null)
                    {
                        foreach (Quote quote in quoteToBeClosed)
                        {
                            if (quote.StateCode == QuoteState.Won)
                            {
                                iWonQuotes++;
                            }
                        }
                    }
                    //if (iWonQuotes == 0)
                    //{
                    //    throw new InvalidPluginExecutionException("Unable to close this opportunity as won. There are no Won quotes.");
                    //}

                    if (quoteToBeClosed != null)
                    {
                        foreach (Quote quote in quoteToBeClosed)
                        {
                            if (quote.StateCode.Value != QuoteState.Won && quote.StateCode.Value != QuoteState.Closed)
                            {
                                if (quote.StateCode.Value == QuoteState.Draft)
                                {
                                    //have to be activated first
                                    ActivateDraftQuote(wService, quote);
                                }

                                //Execute Close as Won
                                ExecuteCloseQuoteAsWon(wService, quote, quote.QuoteNumber, quote.StatusCode.Value);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        // Return all of the Quote from related Opportunity to be Closed
        public List<Quote> GetQuotes(Guid guidOpportunityId)
        {
            Quote quote = new Quote();
            //get the Quote to be canceled in same Opportunity, exclude the wonQuote
            myServiceContext context = new myServiceContext(wService);  //This comes from GeneratedCodeWithContext.cs
            var quoteToBeCanceled = from x in context.QuoteSet
                                    orderby x.CreatedOn descending
                                    where x.OpportunityId.Id == guidOpportunityId
                                    select x;

            if (quoteToBeCanceled.ToList().Count > 0)
            {
                return quoteToBeCanceled.ToList<Quote>();
            }
            else
            {
                return null;
            }
        }

        // Activate the Draft Quote
        public void ActivateDraftQuote(IOrganizationService service, Quote erQuote)
        {

            // Activate the quote
            SetStateRequest activateQuote = new SetStateRequest()
            {
                EntityMoniker = erQuote.ToEntityReference(),
                State = new OptionSetValue((int)QuoteState.Active),
                Status = new OptionSetValue((int)2) //in progress
            };
            service.Execute(activateQuote);
        }

        // Close the Quote as Canceled
        public void ExecuteCloseQuoteAsCanceled(IOrganizationService service, Quote erQuote, string strQuoteNumber, int statusCode)
        {
            int status = (statusCode == 9) ? 7 : 100000003;     // 7 - Expired, 10000003 - Closed Opp Won

            CloseQuoteRequest closeQuoteRequest = new CloseQuoteRequest()
            {
                QuoteClose = new QuoteClose()
                {
                    Subject = String.Format("Quote Closed (Canceled) - {0} - {1}", strQuoteNumber, DateTime.Now.ToString()),
                    QuoteId = erQuote.ToEntityReference()
                },
                Status = new OptionSetValue(status)
            };
            service.Execute(closeQuoteRequest);
        }

        // Close the Quotes as Opportunity Won
        public void ExecuteCloseQuoteAsWon(IOrganizationService service, Quote erQuote, string strQuoteNumber, int statusCode)
        {
            int status = (statusCode == 9) ? 7 : 100000003;     // 7 - Expired, 10000003 - Closed Opp Won
            CloseQuoteRequest closeQuoteRequest = new CloseQuoteRequest()
            {
                QuoteClose = new QuoteClose()
                {
                    Subject = String.Format("Quote Closed (Opportunity Won) - {0} - {1}", strQuoteNumber, DateTime.Now.ToString()),
                    QuoteId = erQuote.ToEntityReference()
                },
                Status = new OptionSetValue(status)  // Opportunity Won
            };
            service.Execute(closeQuoteRequest);
        }
    }
}