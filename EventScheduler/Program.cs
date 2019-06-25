using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;

namespace EventScheduler
{
    public class Program
    {
        public static ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
        public static void Main(string[] args)
        {

            mail();
            createmeeting();
            getmail();
        }

        // Sends an email to the specified user
        public static void mail()
        {
            service.Credentials = new WebCredentials("cstorch@1800contacts.com", "Lymetime1");
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.AutodiscoverUrl("cstorch@1800contacts.com", RedirectionUrlValidationCallback);
            service.UseDefaultCredentials = false;
            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("cstorch@1800contacts.com");
            email.Subject = "Testing C# CODE";
            email.Body = new MessageBody("Testing EWS api to send an email");
            email.Send();
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        // takes data from the emails and creates a calendar appointment with it
        public static void createmeeting(string subject, string body)
        {
            // Create the appointment.
            Appointment appointment = new Appointment(service);

            //Set properties on the appointment.
            appointment.Subject = subject;
            appointment.Body = body;
            appointment.Start = new DateTime(2019, 6, 21, 2, 30, 0);
            appointment.End = appointment.Start.AddHours(2);
            appointment.RequiredAttendees.Add("cstorch@1800contacts.com");

            // Save the appointment.
            appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
        }


        // Pulls the emails and gathers the data needed create a calendar appointment
        public static void getmail()
        {
            // Add a search filter that searches on the body or subject.
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Test"));
            //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Body, "homecoming"));

            // Create the search filter.
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());

            // Create a view with a page size of 50.
            ItemView view = new ItemView(10);

            // Identify the Subject and DateTimeReceived properties to return.
            // Indicate that the base property will be the item identifier
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
            view.PropertySet.Add(ItemSchema.DateTimeSent);


            // Order the search results by the DateTimeReceived in descending order.
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
            view.Traversal = ItemTraversal.Shallow;

            // Send the request to search the Inbox and get the results.
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);
            DateTime fiveMinutesAgo = DateTime.Now.AddMinutes(-5);
            // Process each item.
            foreach (Item myItem in findResults.Items)
            {
                if (myItem.DateTimeReceived >= fiveMinutesAgo)
                {
                    Console.WriteLine("lol");
                    //createmeeting();
                }
            }
        }
    }
}
