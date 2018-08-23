(function(){
  'use strict';

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

    // Obtains the current item.
    var _Item = Office.context.mailbox.item;

    // When forward button is clicked, build the content
    // and send email.
    $('#forward-button').on('click', function(){
      $('#error-display').hide();
      $('#success-display').hide();
      var mailDesc = $('#email-desc').val();
      var docType = $('#doctype-list').find(":selected").text();
      if (mailDesc.length > 0) {
        if (!docType.match("^Choose")) {
          
          postMail(_Item, mailDesc, docType, function(result, error){
            if (error) {
              $('#error-text').html(error);
              $('#error-display').show();
            } else {
              $('#success-text').text(result);
              $('#success-display').show();
            }
          });

        } else {
          showError('Please select a document type.');
        }
      } else {
        showError('Please insert a valid description.');
      }
    });
      
    });
  };

  function showError(error) {
    $('#error-text').text(error);
    $('#error-display').show();
  }

  /*function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected');
    $(this).addClass('is-selected');
    $('#insert-button').removeAttr('disabled');
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }*/
  
})();