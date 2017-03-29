var http = require("https");

var options = {
  "method": "GET",
  "hostname": "raw.githubusercontent.com",
  "port": null,
  "path": "/higherknowledge/outlook-integration/master/templates/srivalli%2540higherknowledge.in",
  "headers": {
    "content-type": "application/json",
    "cache-control": "no-cache",
    "postman-token": "140da308-7068-ff29-82ff-4bc2ec3e39fe"
  }
};

var gitReq = http.gitRequest(options, function (gitRes) {
  var chunks = [];

  gitRes.on("data", function (chunk) {
    chunks.push(chunk);
  });

  gitRes.on("end", function () {
    var body = Buffer.concat(chunks);
    console.log(body.toString());
  });
});

gitReq.write(JSON.stringify({ access_token: '1234' }));
gitReq.end();