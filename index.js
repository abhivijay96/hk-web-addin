var express = require('express');
var app = express();
var http = require('https');

app.set('port', (process.env.PORT || 5000));

app.get('/:name', function (req, res, next) {

  var options = {
    root: __dirname + '/addin/',
    dotfiles: 'deny',
    headers: {
        'x-timestamp': Date.now(),
        'x-sent': true
    }
  };

  var fileName = req.params.name;
  res.sendFile(fileName, options, function (err) {
    if (err) {
      console.log(err);
      res.status(err.status).end();
    }
    else {
      console.log('Sent:', fileName);
    }
  });

});

app.get('/template/:email', function(req, res, next){
  
    if(req.headers["Authorization"] != "hktemplatepass")
    {
        res.statusCode = 403;
        res.end();
    }

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

    var gitReq = http.request(options, function (gitRes) {
      var chunks = [];

      gitRes.on("data", function (chunk) {
        chunks.push(chunk);
      });

      gitRes.on("end", function () {
        var body = Buffer.concat(chunks);
        res.write(body.toString());
        res.end();
      });
    });
    gitReq.end();
  }
  
);

app.listen(app.get('port'), function() {
  console.log('Node app is running on port', app.get('port'));
});
