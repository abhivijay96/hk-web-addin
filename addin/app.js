/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $("#intern").click(handleIntern);
      $("#recruit").click(handleRecruit);
    });
  };

  function run() {

  }

  function handleIntern() {
    fetchTemp(1);
  }



  function handleRecruit() {
    fetchTemp(2);
  }

  function fetchTemp(flag) {

    var req = new XMLHttpRequest();

    function reqListener() {
      if (req.readyState == req.DONE && req.status == 200) {
        send(this.responseText);
        //ga(flag == 1 ? "Intern" : "Full-Time");
      }
    }

    req.onreadystatechange = reqListener;
    req.open("GET", "https://web-addin.herokuapp.com/template/" + Office.context.mailbox.userProfile.emailAddress.toLowerCase() + (flag == 1 ? "" : "R"));
    req.setRequestHeader("Authorization", "hktemplatepass");
    req.send();
  }

  function ga(eve) {
    function listener() {
      return;
    }

    var req = new XMLHttpRequest();
    req.onreadystatechange = listener;
    req.open("POST", "https://www.google-analytics.com/collect");
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

  function getBody(body) {
    var res = "";
    body.forEach(function (entry) {
      res += entry + "<br/><br/>";
    })
    return res;
  }

})();