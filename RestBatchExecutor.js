
var RestBatchExecutor = (function () {
    function RestBatchExecutor(appWebUrl, authHeader) {
        this.changeRequests = [];
        this.getRequests = [];
        this.resultsIndex = [];
        this.appWebUrl = appWebUrl;
        this.authHeader = authHeader;
    }
    RestBatchExecutor.prototype.loadChangeRequest = function (request) {
        request.resultToken = this.getUniqueId();
        this.changeRequests.push(request);
        return request.resultToken;
    };

    RestBatchExecutor.prototype.loadRequest = function (request) {
        request.resultToken = this.getUniqueId();
        this.getRequests.push(request);
        return request.resultToken;
    };

    RestBatchExecutor.prototype.executeAsync = function (info, success, error) {
        var dfd = $.Deferred();
        var payload = this.buildBatch();

        if (info && info.crossdomain === true) {
            this.executeCrossDomainAsync(payload).done(function (result) {
                dfd.resolve(result);
            }).fail(function (err) {
                dfd.reject(err);
            });
        } else {
            this.executeJQueryAsync(payload).done(function (result) {
                dfd.resolve(result);
            }).fail(function (err) {
                dfd.reject(err);
            });
        }

        return dfd;
    };

    RestBatchExecutor.prototype.executeCrossDomainAsync = function (batchBody) {
        var _this = this;
        var dfd = $.Deferred();
        var batchUrl = this.appWebUrl + "/_api/$batch";
        var executor = new SP.RequestExecutor(this.appWebUrl);
        var info = {
            url: batchUrl,
            method: "POST",
            body: batchBody,
            headers: this.getRequestHeaders(),
            success: function (data) {
                var results = _this.buildResults(data.body);
                _this.clearRequests();
                dfd.resolve(results);
            },
            error: function (err) {
                _this.clearRequests();
                dfd.reject(err);
            }
        };

        executor.executeAsync(info);
        return dfd;
    };

    RestBatchExecutor.prototype.executeJQueryAsync = function (batchBody) {
        var _this = this;
        var dfd = $.Deferred();
        var batchUrl = this.appWebUrl + "/_api/$batch";

        $.ajax({
            'url': batchUrl,
            'type': 'POST',
            'data': batchBody,
            'headers': this.getRequestHeaders(),
            'success': function (data) {
                var results = _this.buildResults(data);
                _this.clearRequests();
                dfd.resolve(results);
            },
            'error': function (err) {
                _this.clearRequests();
                dfd.reject(err);
            }
        });

        return dfd;
    };

    RestBatchExecutor.prototype.getRequestHeaders = function () {
        var header = {};
        header['accept'] = 'application/json;odata=verbos';
        header['content-type'] = 'multipart/mixed; boundary=batch_8890ae8a-f656-475b-a47b-d46e194fa574';
        header[Object.keys(this.authHeader)[0]] = this.authHeader[Object.keys(this.authHeader)[0]];

        return header;
    };

    RestBatchExecutor.prototype.getBatchRequestHeaders = function (headers, batchCommand) {
        var isAccept = false;
        if (headers) {
            $.each(Object.keys(headers), function (k, v) {
                batchCommand.push(v + ": " + headers[v]);
                if (!isAccept) {
                    isAccept = (v.toUpperCase() === "ACCEPT");
                }
                ;
            });
        }

        if (!isAccept) {
            batchCommand.push('accept:application/json;odata=verbose');
        }
    };

    RestBatchExecutor.prototype.buildBatch = function () {
        var _this = this;
        var batchCommand = [];
        var batchBody;

        $.each(this.changeRequests, function (k, v) {
            _this.buildBatchChangeRequest(batchCommand, v, k);
            _this.resultsIndex.push(v.resultToken);
        });

        batchCommand.push("--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e--");

        $.each(this.getRequests, function (k, v) {
            _this.buildBatchGetRequest(batchCommand, v, k);
            _this.resultsIndex.push(v.resultToken);
        });

        batchBody = batchCommand.join('\r\n');

        //embed all requests into one batch
        batchCommand = new Array();
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574");
        batchCommand.push("Content-Type: multipart/mixed; boundary=changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e");
        batchCommand.push('Content-Length: ' + batchBody.length);
        batchCommand.push('Content-Transfer-Encoding: binary');
        batchCommand.push('');
        batchCommand.push(batchBody);
        batchCommand.push('');
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574--");

        batchBody = batchCommand.join('\r\n');
        return batchBody;
    };

    RestBatchExecutor.prototype.buildBatchChangeRequest = function (batchCommand, request, batchIndex) {
        batchCommand.push("--changeset_f9c96a07-641a-4897-90ed-d285d2dbfc2e");
        batchCommand.push("Content-Type: application/http");
        batchCommand.push("Content-Transfer-Encoding: binary");
        batchCommand.push("Content-ID: " + (batchIndex + 1));
        batchCommand.push(request.binary ? "processData: false" : "processData: true");
        batchCommand.push('');
        batchCommand.push(request.verb.toUpperCase() + " " + request.endpoint + " HTTP/1.1");
        this.getBatchRequestHeaders(request.headers, batchCommand);
        if (!request.binary && request.payload) {
            batchCommand.push("Content-Type: application/json;odata=verbose");
        }
        if (request.binary && request.payload) {
            batchCommand.push("Content-Length :" + request.payload.byteLength);
        }
        batchCommand.push('');

        if (request.payload) {
            batchCommand.push(request.binary ? request.payload : JSON.stringify(request.payload));
            batchCommand.push('');
        }
    };

    RestBatchExecutor.prototype.buildBatchGetRequest = function (batchCommand, request, batchIndex) {
        batchCommand.push("--batch_8890ae8a-f656-475b-a47b-d46e194fa574");
        batchCommand.push('Content-Type: application/http');
        batchCommand.push('Content-Transfer-Encoding: binary');
        batchCommand.push("Content-ID: " + (batchIndex + 1));
        batchCommand.push('');
        batchCommand.push('GET ' + request.endpoint + ' HTTP/1.1');
        this.getBatchRequestHeaders(request.headers, batchCommand);
        batchCommand.push('');
    };

    RestBatchExecutor.prototype.buildResults = function (responseBody) {
        var _this = this;
        var responseBoundary = responseBody.substring(0, 52);
        var resultTemp = responseBody.split(responseBoundary);
        var resultData = [];

        $.each(resultTemp, function (k, v) {
            if (v.indexOf('\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary') == 0) {
                var responseTemp = v.split('\r\n');
                var batchResult = new RestBatchResult();

                //grab just the http status code
                batchResult.status = responseTemp[4].substr(9, 3);

                //based on the status pull the result from response
                batchResult.result = _this.getResult(batchResult.status, responseTemp);

                //assign return token to result
                resultData.push({ id: _this.resultsIndex[k - 1], result: batchResult });
            }
        });

        return resultData;
    };

    RestBatchExecutor.prototype.getResult = function (status, response) {
        switch (status) {
            case "400":
            case "404":
            case "500":
            case "200":
                return this.parseJSON(response[7]);
            case "204":
            case "201":
                return this.parseJSON(response[9]);
            default:
                return this.parseJSON(response[4]);
        }
    };

    RestBatchExecutor.prototype.getUniqueId = function () {
        return (this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum() + this.randomNum());
    };

    RestBatchExecutor.prototype.randomNum = function () {
        return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
    };

    RestBatchExecutor.prototype.clearRequests = function () {
        while (this.changeRequests.length) {
            this.changeRequests.pop();
        }
        ;
        while (this.getRequests.length) {
            this.getRequests.pop();
        }
        ;
        while (this.resultsIndex.length) {
            this.resultsIndex.pop();
        }
        ;
    };

    RestBatchExecutor.prototype.parseJSON = function (jsonString) {
        try  {
            var o = JSON.parse(jsonString);

            // Handle non-exception-throwing cases:
            // Neither JSON.parse(false) or JSON.parse(1234) throw errors, hence the type-checking,
            // but... JSON.parse(null) returns 'null', and typeof null === "object",
            // so we must check for that, too.
            if (o && typeof o === "object" && o !== null) {
                return o;
            }
        } catch (e) {
        }

        return jsonString;
    };
    return RestBatchExecutor;
})();

var RestBatchResult = (function () {
    function RestBatchResult() {
        this.status = "";
        this.result = null;
    }
    return RestBatchResult;
})();
var BatchRequest = (function () {
    function BatchRequest() {
        this.resultToken = "";
        this.endpoint = "";
        this.payload = "";
        this.binary = false;
        this.headers = null;
        this.verb = "GET";
    }
    return BatchRequest;
})();
