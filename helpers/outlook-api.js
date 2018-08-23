// Initialize a context object for the add-in.
// Set the fields that are used on the request
// object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };

function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}

function postMail(_Item, descrPost, doctypePost, callback) {
  var requestUrl = 'https://localhost:8443/entities.mail';

  var attachments = [];

  if (_Item.attachments.length > 0) {
    for (i = 0 ; i < _Item.attachments.length ; i++) {
      var _att = _Item.attachments[i];
      var attObj = { 
        "attchId": _att.id, 
        "fileContents": _att.contentType, 
        "name": _att.name };
      attachments.push(attObj);
    }
  }

  _Item.body.getAsync(
    "text",
    { asyncContext:"This is passed to the callback" },
    function asyncallback(result) {

      // json request
      var _json = JSON.stringify({
        itemId: _Item.itemId,
        descr: descrPost,
        doctype: doctypePost,
        subject: _Item.subject,
        contents: result.value, //result.value contains mail contents
        attachmentsList: attachments
      });

      // Send the data using post
      var jqxhr = $.ajax({
        url: requestUrl,
        contentType: "application/json; charset=utf-8",
        type: 'POST',
        data: _json
      });

      jqxhr.done(function( data ) {
        callback(data);
      });

      jqxhr.fail(function( errorJqXHR, textStatus, errorThrown ) {
        var err = parseError (errorJqXHR, textStatus, errorThrown);
        callback(null, err);
      });      
    }
  );
}

function parseError (errorJqXHR, textStatus, errorThrown) {
  var details = '<div style="white-space: normal;"><strong>status code:</strong> ' + errorJqXHR.status + '<br>'
                + '<strong>textStatus:</strong> ' + textStatus + '<br>'
                + '<strong>errorThrown:</strong> ' + errorThrown + '<br>'
                + '<strong>errorJqXHR.responseText:</strong><br>' 
                + '<div>' + errorJqXHR.responseText + '</div><br>'
                + '<i>Look at the browser console (F12 or Ctrl+Shift+I, Console tab) for more information!</i></div>';
  return details;
}