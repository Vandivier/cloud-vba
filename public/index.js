let cloud = {                                                                         // could be a seperate module
  processStorageEvent: function(e) {
    let oEvent = JSON.parse(e.newValue)

    if (e.key === 'cloud-message') {
      switch(oEvent.action) {
        case 'userresponse':
            $('#response-region').text(decodeURIComponent(oEvent.options.data))
            break;
        default:
            // n/a
      }
    }

  }
}

$(function(){

    $('#trigger, #authorize').click(function(){
      let wCloud;
      let sLocation;
      let sOptions;

      sLocation = 'ms-word:nft|u|'                                                    // prefix required to open Word by URL. More info in the readme.
      sLocation += window.location.protocol + window.location.host                    // make it environmant-agnostic
      sLocation += '/files/cloud.dotm?action='                                        // call the macro template and pass message by querystring

      if ($(this)[0].id === 'trigger') {
        sOptions = 'width=100,height=100,top=1000,left=1000'                          // try to obfuscate the popup. Want to avoid it? You should be posting, not using a dotm.
        sLocation += 'PromptUserForInput'
      } else {
        sLocation += 'JustDie'                                                        // call the macro template and pass message by querystring
      }
      wCloud = window.open(sLocation, '_blank', sOptions)
      setTimeout(function(){ wCloud.close() }, 250)

    });

    window.addEventListener('storage',  (e) => cloud.processStorageEvent(e));

});
