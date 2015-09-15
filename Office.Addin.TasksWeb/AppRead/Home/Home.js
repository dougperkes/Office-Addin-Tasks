/// <reference path="../App.js" />
/// <reference path="../../Scripts/jquery-1.9.1.js" />

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
		$("#taskSubject").val(item.subject);
		$("#insertTaskBody").click(function() {
			item.body.getAsync("text", function(result) {
				$("#taskBody").val(result.value);
			});
		});
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
	
	/*
	FindItem Operation: https://msdn.microsoft.com/EN-US/library/aa566107(v=exchg.150).aspx
	QueryString: https://msdn.microsoft.com/EN-US/library/ee693615(v=exchg.150).aspx
	*/
	function findRelatedTasks() {
		var soapEnv = '<?xml version="1.0" encoding="utf-8"?> \
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" \
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> \
  <soap:Body> \
    <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" \
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" \
              Traversal="Shallow"> \
		<m:ItemShape> \
        	<t:BaseShape>IdOnly</t:BaseShape> \
        	<t:AdditionalProperties> \
          		<t:FieldURI FieldURI="item:Subject" /> \
        	</t:AdditionalProperties> \
      	</m:ItemShape> \
      <m:IndexedPageItemView MaxEntriesReturned="1" Offset="0" BasePoint="Beginning" /> \
      <m:QueryString>Kind:tasks</m:QueryString> \
    </FindItem> \
  </soap:Body> \
</soap:Envelope>';

		var mailbox = Office.context.mailbox;
		mailbox.makeEwsRequestAsync(soapEnv, searchCallback);
	}
	
	function searchCallback(asyncResult) {
		var result = asyncResult.value;
		var context = asyncResult.context;
		
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
								  <t:Subject>' + $("#taskSubject").val() + '</t:Subject> \
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
	   mailbox.makeEwsRequestAsync(soapEnv, createTaskCallback);
	}

function createTaskCallback(asyncResult)  {
	app.showNotification('Status', 'EWS Request Completed');
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
   console.log(result);
   console.log(context);

}

		
	
})();