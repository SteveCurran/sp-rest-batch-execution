# sp-rest-batch-execution
Javascript library for executing SharePoint REST commands in batches for O365 only. To be used with O365 new REST batch requests.
#Example
The RestBatchExecutor encapsulates all the complexity of wrapping your REST requests into one change set and batch. First create a new RestBatchExecutor. The constructor requires the O365 application web URL and an authentication header. The URL will be used to construct the $Batch endpoint where the requests will be submitted. The authentication header in the form of a JSON object allows for you to either use the formDigest or the OAuth token. 

The next step is to create a new BatchRequest for each request to be batched. Set the BatchRequest’s endpoint property to your REST endpoint. Second set the payload property to any JSON object you want to send with your request, this is typically what you would put in the data property of an JQuery $ajax request.  Third, set the verb property. The verb property represents the HTTP request you typically use. For example, if you are updating a list item then use the verb MERGE. This is always set using the “X-HTTP-Method” header. However this verb must be used at the beginning of your endpoint when submitting requests to $Batch. Other verbs would be POST,PUT,DELETE. Finally you can optionally set the headers property. In the case of a DELETE, MERGE or PUT you should set your “If-Match” header to either the etag of the entity or an “*”.  The headers also allows you to take advantage of JSON Light by setting the “accept” header to “application/json;odata=nometadata” for example. 
The example below shows three defined endpoints and the creation of three batch requests, representing a list item creation, update and retrieval of the list. After creating a BatchRequest you will need to add it to the RestBatchExecutor using either the loadChangeRequest or loadRequest method. The loadChangeRequest should only be used to add requests that use the POST,DELETE,MERGE or PUT verbs. This makes sure all your write requests are sent in one change request. Use the loadRequest method when doing any type of GET requests. always save the unique token that is returned by both these methods. This token will be used to access the results. In the example I assign the token to an array along with a title for the operation. 

```JavaScript
var createEndPoint = appweburl
    + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('coolwork')/items?@target='" + hostweburl + "'";

var updateEndPoint = appweburl
    + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('coolwork')/items(134)?@target='" + hostweburl + "'";

var getEndPoint = appweburl
    + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('coolwork')/items?@target='" + hostweburl + "'&$orderby=Title";

var commands = [];

batchRequest = new BatchRequest();
batchRequest.endpoint = createEndPoint;
batchRequest.payload = { '__metadata': { 'type': 'SP.Data.CoolworkListItem' }, 'Title': 'SharePoint REST' };
batchRequest.verb = "POST"
commands.push({ id: batchExecutor.loadChangeRequest(batchRequest), title: 'Rest Batch Create' });

var batchRequest = new BatchRequest();
batchRequest.endpoint = updateEndPoint;
batchRequest.payload = { '__metadata': { 'type': 'SP.Data.CoolworkListItem' }, 'Title': 'O365 REST' };
batchRequest.headers = { 'IF-MATCH': "*" };
batchRequest.verb = "MERGE";
commands.push({ id: batchExecutor.loadChangeRequest(batchRequest), title: 'Rest Batch Update' });

batchRequest = new BatchRequest();
batchRequest.endpoint = getEndPoint;
batchRequest.headers = { 'accept': 'application/json;odata=nometadata' }
commands.push({ id: batchExecutor.loadRequest(batchRequest), title: "Rest Batch Get Items" });

batchExecutor.executeAsync().done(function (result) {
    var d = result;
    var msg = [];
    $.each(result, function (k, v) {
        var command = $.grep(commands, function (command) {
            return v.id === command.id;
        });
        if (command.length) {
            msg.push("Command--" + command[0].title + "--" + v.result.status);
        }
    });

    alert(msg.join('\r\n'));

}).fail(function (err) {
    alert(JSON.stringify(err));
});
```
