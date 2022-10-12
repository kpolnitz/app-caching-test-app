import React, { ChangeEvent } from "react";
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

  const onBtnChange = (event: ChangeEvent<HTMLInputElement>) => {
    console.log("we hit the button in config");
  };

  return (
    <div>
      <h1>Tab Configuration</h1>
      <div>
        This is where you will add your tab configuration options the user can
        choose when the tab is added to your team/group chat.
      </div>
      <input
        type="radio"
        name="TabConfig"
        onChange={onBtnChange}
      />{" "}
      Normal
    </div>
  );
};
