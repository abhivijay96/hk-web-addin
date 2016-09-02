
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
      }

      var oReq = new XMLHttpRequest();
      oReq.addEventListener("load", reqListener);
      oReq.open("GET", "https://raw.githubusercontent.com/higherknowledge/outlook-integration/master/templates/" + Office.context.mailbox.userProfile.emailAddress + (flag == 1 ? "" : "R"));
      oReq.send();
  }

  function send(template) {
      var response = JSON.parse(template);
      var body = getBody(response["Body"]);
      Office.context.mailbox.item.displayReplyForm(body);
      localStorage.clear();
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