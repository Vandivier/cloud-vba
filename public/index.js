$(function(){

    $('#trigger').click(function(){                                 // send request to word utilitizing the querystring
      let wCloud;
      let sLocation;
      let sOptions;

      sLocation = 'ms-word:nft|u|'                                  // prefix required to open Word by URL. More info in the readme.
      sLocation += window.location.protocol + window.location.host  // make it environmant-agnostic
      sLocation += '/files/cloud.dotm?action=PromptUserForInput'    // call the macro template and pass message by querystring

      sOptions = 'width=100,height=100,top=1000,left=1000'          // try to obfuscate the popup. Want to avoid it? You should be posting, not using a dotm.
      wCloud = window.open(sLocation, '_blank', sOptions)
      wCloud.close();

    });

});
