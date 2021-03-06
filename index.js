var express = require('express');
var app = express();
var http = require('https');
var qs = require("querystring");

app.set('port', (process.env.PORT || 5000));

app.options('/*', function(req, res, next)
{
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  res.end();
}
);

app.use(express.static('addin'));
// app.get('/:name', function (req, res, next) {

//   var options = {
//     root: __dirname + '/addin/',
//     dotfiles: 'deny',
//     headers: {
//         'x-timestamp': Date.now(),
//         'x-sent': true
//     }
//   };

//   var fileName = req.params.name;
//   res.sendFile(fileName, options, function (err) {
//     if (err) {
//       console.log(err);
//       res.status(err.status).end();
//     }
//     else {
//       console.log('Sent:', fileName);
//     }
//   });

// });

function sendGoogleAnalytics(email, type)
{
  var options = {
    "method": "POST",
    "hostname": "www.google-analytics.com",
    "port": null,
    "path": "/collect",
    "headers": {
      "content-type": "application/x-www-form-urlencoded"
    }
  };
  
  var req = http.request(options, function (res) {
    var chunks = [];
  
    res.on("data", function (chunk) {
      chunks.push(chunk);
    });
  
    res.on("end", function () {
      var body = Buffer.concat(chunks);
      console.log(body.toString());
    });
  });
  
  req.write(qs.stringify({ v: '1',
    t: 'event',
    tid: 'UA-81367328-1',
    cid: '1',
    ec: type,
    el: "fetched",
    ea: email }));
  req.end();
}

app.get('/template/:email', function(req, res, next){
  
    if(req.get("Authorization").toString() != "hktemplatepass")
    {
        res.statusCode = 401;
        res.end();
        return;
    }
    
    console.log("/higherknowledge/outlook-integration/master/templates/" + req.params.email);
    var options = {
    "method": "GET",
    "hostname": "raw.githubusercontent.com",
    "port": null,
    "path": "/higherknowledge/outlook-integration/master/templates/" + req.params.email,
    "headers": {
      "content-type": "application/json",
      "cache-control": "no-cache"
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
        sendGoogleAnalytics(req.params.email, 'used addin');
      });
    });
    gitReq.end();
  }

);

app.listen(app.get('port'), function() {
  console.log('Node app is running on port', app.get('port'));
});
