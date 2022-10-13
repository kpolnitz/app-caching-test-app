import React, { useState } from "react";
import "./App.css";
import * as microsoftTeams from "@microsoft/teams-js";

function logItem(action: string, actionColor: string, message: string) {
  const newItem =
    "<span style='font-weight:bold;color:" +
    actionColor +
    "'>" +
    action +
    "</span> " +
    message +
    "</br>";
  return newItem;
}

const beforeUnloadHandler = (
  setItems: React.Dispatch<React.SetStateAction<string[]>>,
  readyToUnload: () => void
) => {
  console.log("got beforeUnload from TEAMS");
  let newItem = logItem("OnBeforeUnload", "purple", "Started");
  setItems((oldItems) => [...oldItems, newItem]);

  setTimeout(() => {
    newItem = logItem("OnBeforeUnload", "purple", "Completed");
    setItems((oldItems) => [...oldItems, newItem]);
    console.log("sending readyToUnload to TEAMS");
    readyToUnload();
  }, 2000);
  return true;
};

const loadHandler = (
  setItems: React.Dispatch<React.SetStateAction<string[]>>,
  data: microsoftTeams.LoadContext
  ) => {
  console.log("got load from TEAMS", data);
  logItem("OnLoad", "blue", "Started for " + data.entityId);

  const timeout = 1000;
  setTimeout(() => {
    let newItem = logItem("OnLoad", "blue", "Completed for " + data.entityId);
    setItems((oldItems) => [...oldItems, newItem]);
    console.log("sending notifyAppLoaded to TEAMS");

    microsoftTeams.app.notifySuccess();
  }, timeout);
};

export const Tab = () => {
  const [items, setItems] = useState<string[]>([]);
  const [title, setTitle] = useState("initial title");
  const [initState, setInitState] = useState(false);
  React.useEffect(() => {
    if (!initState) {
      return;
    }
    window.performance.mark("Teams-GetTabContextStart");
    // get context


    microsoftTeams.app.getContext().then(
      (context: microsoftTeams.app.Context) => {
      window.performance.mark("Teams-GetTabContextEnd");
      window.performance.measure(
        "Teams-GetTabContext",
        "Teams-GetTabContextStart",
        "Teams-GetTabContextEnd"
      );
      if (context) {
        console.log("got context from TEAMS", context);
        const newItem = logItem("Success", "green", "Loaded Teams context");
        setItems((oldItems) => [...oldItems, newItem]);
        setTitle(context.page.id);

        if (context.page.frameContext === "sidePanel") {
          // ############################################
          // OnBeforeUnload
          microsoftTeams.teamsCore.registerBeforeUnloadHandler((readyToUnload) => {
            const result = beforeUnloadHandler(setItems, readyToUnload);
            return result;
          });

          // ############################################
          // OnLoad
          microsoftTeams.teamsCore.registerOnLoadHandler((data) => {
            loadHandler(setItems, data);
          });          
        }
      } else {
        let newItem = logItem("ERROR", "red", "could not get context");
        setItems((oldItems) => [...oldItems, newItem]);
      }
    });
    return () => {
      console.log("useEffect cleanup - Tab");
    };
  }, [initState]);

  React.useEffect(() => {
    const timeout = 2000;
    setTimeout(() => {
      console.log("sending notifySuccess to TEAMS");
      microsoftTeams.app.notifySuccess();
      setInitState(true);
    }, timeout);
  }, []);

  React.useEffect(() => {
    if (initState) {
      console.log("invoke auth token");
      try {
        microsoftTeams.authentication.getAuthToken()
      }
      catch {
        console.log("error");
      }
    }
  }, [initState]);

  const jsx = initState ? (
    <div>
      <h3>Entity ID - {title}</h3>
      {items.map((item) => {
        return <div dangerouslySetInnerHTML={{ __html: item }} />;
      })}
    </div>
  ) : (
    <div style={{ color: "white" }}>loading</div>
  );
  return jsx;
};
