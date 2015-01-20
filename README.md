# sp-rest-batch-execution
Javascript library for executing SharePoint REST commands in batches for O365 only. To be used with O365 new REST batch requests.
#Example
Create a RestBatchExecutor using your app web URL. You can then create a BatchRequest. Use the BatchRequest's endpoint property to set the typical REST endpoint. Use the payload property to set the JSON data when posting. Call the RestBatchExecutor.loadChangeRequest method with the BatchRequest, then assign the returned id to an array to retrieve the result. When you just want to retrieve (GET) just set the BatchRequest's endpoint property and load the BatchRequest using the RestBatchExecutor's loadRequest method. Finally call the RestBatchExecutor's executeAsync method. The results will be an array of key values. The key represents the Id that was returned when loading the batch and the value represents a BatchResult. The BatchResult has a status property containing a typical HTTP status like 201 or 400, and a result property containing the JSON result. 


    var createEndPoint = appweburl
                       + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('coolwork')/items?@target='" + hostweburl + "'";
    var getEndPoint = appweburl
                    + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('coolwork')/items?@target='" + hostweburl +   "'&$orderby=Title";
    var commands = [];
    var batchExecutor = new RestBatchExecutor(appweburl);
    var batchRequest = new BatchRequest();
    batchRequest.endPoint = createEndPoint;
    batchRequest.payload = { '__metadata': { 'type': 'SP.Data.CoolworkListItem' }, 'Title': 'item1'};
    commands.push({ id: batchExecutor.loadChangeRequest(batchRequest), title: 'item1' });

    batchRequest = new BatchRequest();
    batchRequest.endPoint = createEndPoint;
    batchRequest.payload = { '__metadata': { 'type': 'SP.Data.CoolworkListItem' }, 'Title': 'item2' };
    commands.push({ id: batchExecutor.loadChangeRequest(batchRequest), title: 'item2' });

    batchRequest = new BatchRequest();
    batchRequest.endPoint = getEndPoint;
    commands.push({ id: batchExecutor.loadRequest(batchRequest), title: "get coolwork items" });
    
    batchExecutor.executeAsync().done(function (result) {
        var d = result;
        $.each(result, function (k, v) {
            alert(v.result.status + "--" + $.grep(commands, function (command) {
                return v.id === command.id;
            })[0].title);
        });

    }).fail(function (err) {
        alert(JSON.stringify(err));
    });
