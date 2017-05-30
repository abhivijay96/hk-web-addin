
(function () {

  var messageBanner;
  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();

      let debug = document.getElementById("debug");
      debug.innerHTML = ""
      window.onerror = function (msg, url, lineNo, columnNo, error) {
            var string = msg.toLowerCase();
            var substring = "script error";
            if (string.indexOf(substring) > -1){
                //alert('Script Error: See Browser Console for Detail');
                debug.innerHTML = "Script Error";
            } 
            else {
                var message = [
                    'Message: ' + msg,
                    'URL: ' + url,
                    'Line: ' + lineNo,
                    'Column: ' + columnNo,
                    'Error object: ' + JSON.stringify(error)
                ].join(' - ');
                debug.innerHTML = JSON.stringify(message);
            }

            return false;
        };
    });
  };


  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
      var item = Office.context.mailbox.item;
      var address = Office.context.mailbox.userProfile.emailAddress;
    
    function handleIntern()
      {
        //if(localStorage["intern"] == undefined && !localStorage["hasIntern"])
        {
            fetchTemp(1);
        }

        // else
        // {
        //     send(localStorage["intern"]);
        // }      
    }

    

    function handleRecruit()
    {
        console.log("clicked recruit");
        //if(localStorage["recruit"] == undefined && !localStorage["hasRecruit"])
        {
            fetchTemp(2);
        }
        // else
        // {
        //     send(localStorage["recruit"]);
        //     //ga("sent");
        // }
    }

      $("#intern").click(handleIntern);
      $("#recruit").click(handleRecruit);

  }

  function fetchTemp(flag) {

      var req = new XMLHttpRequest();

      function reqListener() {
          if(req.readyState == req.DONE && req.status == 200)
          {
              if (flag == 1)
               {
                     localStorage["intern"] = this.responseText;
                     localStorage["hasIntern"] = true;
               }
              else
               {
                     localStorage["recruit"] = this.responseText;
                     localStorage["hasRecruit"] = true;
               }
                send(this.responseText);
                //   ga("sent");
                var day = new Date();
                localStorage["fetched"] = day; 
          }
      }

      req.onreadystatechange = reqListener;
      req.open("GET", "https://web-addin.herokuapp.com/template/" + Office.context.mailbox.userProfile.emailAddress.toLowerCase() + (flag == 1 ? "" : "R"));
      req.setRequestHeader("Authorization", "hktemplatepass");
      req.send();
  }

  function ga(eve){
      function listener(){
          return;
      }

      var req = new XMLHttpRequest();
      req.onreadystatechange = listener;
      req.open("POST","https://www.google-analytics.com/collect");
      var data = "v=1&t=event&tid=UA-81367328-1&cid=1";
       data += "&ec=" + Office.context.mailbox.userProfile.emailAddress + "&el=Used Add in" + "&ev=1";
       data += "&ea=" + eve;
       req.send(data);
  }

  function send(template) {
      var response = JSON.parse(template);
      var body = getBody(response["Body"]);
      var reply = Office.context.mailbox.item.displayReplyForm(body);
  }

  function getBody(body)
  {
      var res = "";
      body.forEach(function (entry) {
          res += entry + "<br/><br/>";
      })
      return res;
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();
