
(function () {
  "use strict";

  var messageBanner;
  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  };


  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
      var item = Office.context.mailbox.item;
      var address = Office.context.mailbox.userProfile.emailAddress;
     
      $("#intern").click(handleIntern);
      $("#recruit").click(handleRecruit);
     
    
    function handleIntern()
      {
        
        if(localStorage["intern"] == null)
        {
            fetchTemp(1);
        }
        else
        {
            send(localStorage["intern"]);
            ga("sent");
        }
        
    }

    

    function handleRecruit()
    {
        if(localStorage["recruit"] == null)
        {
            fetchTemp(2);
        }
        else
        {
            send(localStorage["recruit"]);
            ga("sent");
        }
        
    }

  }

  function fetchTemp(flag) {
      function reqListener() {
          if (flag == 1)
              localStorage["intern"] = this.responseText;
          else
              localStorage["recruit"] = this.responseText;
          send(this.responseText);
          ga("sent");
      }

      var oReq = new XMLHttpRequest();
      oReq.addEventListener("load", reqListener);
      oReq.open("GET", "https://raw.githubusercontent.com/higherknowledge/outlook-integration/master/templates/" + Office.context.mailbox.userProfile.emailAddress + (flag == 1 ? "" : "R"));
      oReq.send();
  }

  function ga(eve){
      function listener(){
          showNotification("RES",this.responseText);
      }

      var req = new XMLHttpRequest();
      req.addEventListener("load",listener);
      req.open("POST","https://www.google-analytics.com/collect");
      let data = "v=1&t=event&tid=UA-81367328-1&cid=1";
       data += "&ec=" + Office.mailbox.userProfile.emailAddress + "&el=Used Add in" + "&ev=1";
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
