
(function ($) {
    // Just a random helper to prettify json
    function prettify(data) {
      var json = JSON.stringify(data);
      json = json.replace(/{/g, "{\n\n\t");
      json = json.replace(/,/g, ",\n\t");
      json = json.replace(/}/g, ",\n\n}");
      return json;
    }
  
    $("document").ready(function () {
      // Get the output pre
      var output = $("#output");

      // Determine if we are running inside of an authentication dialog
      // If so then just terminate the running function
      if (OfficeHelpers.Authenticator.isAuthDialog()) {
        // Adding code here isn"t guaranteed to run as we need to close the dialog
        // Currently we have no realistic way of determining when the dialog is completely
        // closed.
        output.text("Closing AuthDialog...");
        return;
      }
  
      // Create a new instance of Authenticator
      var authenticator = new OfficeHelpers.Authenticator();
      
      // Register our providers accordingly
      authenticator.endpoints.registerAzureADAuth("CLIENT_ID_GUID_HERE", "TENANT_ID_GUID_HERE", {
        redirectUrl: 'https://your_domain.com/index.html',
        resource: 'https://graph.microsoft.com'
      });
  

      // Authenticate with the chosen provider
      authenticator.authenticate(OfficeHelpers.DefaultEndpoints.AzureAD, true /* setting the force to true, always re-authenticates. This is just for demo purposes */)
        .then(function (token) {
          // Consume the acess token
          output.text(prettify(token));
        })
        .catch(function (error) {
          // Handle the error
          output.text(prettify(error));
        });

      // Add event handlers to the buttons
      $(".login").click(function () {
        var token = authenticator.tokens.get(OfficeHelpers.DefaultEndpoints.AzureAD);
        output.text("Token Info: " + prettify(token));
      });
    });
  })(jQuery);
