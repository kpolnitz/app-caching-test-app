import React from "react";
import "./App.css";
import * as microsoftTeams from "@microsoft/teams-js";

export const TabConfig = () => {
  microsoftTeams.appInitialization.notifySuccess();
  /**
   * When the user clicks "Save", save the url for your configured tab.
   * This allows for the addition of query string parameters based on
   * the settings selected by the user.
   */
  React.useEffect(() => {
    microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
      const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
      var entityId = "AppInstance_" + Math.floor(Math.random() * 100 + 1);
        microsoftTeams.settings.setSettings({
          suggestedDisplayName: "Tab",
          entityId: entityId,
          contentUrl: baseUrl + `/tab`,
          websiteUrl: baseUrl + `/tab`,
        });
      saveEvent.notifySuccess();
    });
  }, []);
  /**
   * After verifying that the settings for your tab are correctly
   * filled in by the user you need to set the state of the dialog
   * to be valid.  This will enable the save button in the configuration
   * dialog.
   */
  microsoftTeams.settings.setValidityState(true);

  return (
    <div>
      <h1>App Caching</h1>
      <div>
        This is the test app for app caching. This app only works in the side panel for testing purposes.
      </div>
    </div>
  );
};
