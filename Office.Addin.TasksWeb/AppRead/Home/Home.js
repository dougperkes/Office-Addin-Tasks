/// <reference path="../App.js" />

(function () {
    "use strict";

	var propertySetId = "d7abdb30-e17e-43ab-8879-f42f7b5efa03";
	var mailMessageId;
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
			$("#createTask").click(createSampleTask);
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
			//"<BLU202-W753EAC796EB01F36FDA6CC6780@phx.gbl>"
			mailMessageId = xmlEscape(item.internetMessageId);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }

	function xmlEscape(val) {
		val = val.replace("<", "&amp;lt");
		val = val.replace(">", "&amp;gt");
		return val;
	}

	function createSampleTask() {
		var soapEnv = '<?xml version="1.0" encoding="utf-8"?> \
						<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" \
									   xmlns:xsd="http://www.w3.org/2001/XMLSchema" \
									   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" \
									   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> \
						  <soap:Body> \
							<CreateItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" \
										xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" \
										MessageDisposition="SaveOnly"> \
							  <Items> \
								<t:Task> \
								  <t:Subject>My task</t:Subject> \
								  <t:ExtendedProperty> \
									<t:ExtendedFieldURI PropertySetId="' + propertySetId + '" \
										PropertyName="RelatedMailMessage" PropertyType="String" /> \
									<t:Value>' + mailMessageId + '</t:Value> \
								</t:ExtendedProperty> \
								  <t:DueDate>2006-10-26T21:32:52</t:DueDate> \
								  <t:Status>NotStarted</t:Status> \
								</t:Task> \
							  </Items> \
							</CreateItem> \
						  </soap:Body> \
						</soap:Envelope>';
		var mailbox = Office.context.mailbox;
		app.showNotification('Status', 'Making EWS request');
	   mailbox.makeEwsRequestAsync(soapEnv, callback);
	}

function callback(asyncResult)  {
	app.showNotification('Status', 'EWS Request Completed');
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
   console.log(result);
   console.log(context);

}

		
	
})();