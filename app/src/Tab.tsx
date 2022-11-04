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
  newItem = logItem("OnBeforeUnload", "purple", "Completed");
  setItems((oldItems) => [...oldItems, newItem]);
  console.log("sending readyToUnload to TEAMS");
  readyToUnload();
  return true;
};

const loadHandler = (
  setItems: React.Dispatch<React.SetStateAction<string[]>>,
  data: microsoftTeams.LoadContext
  ) => {
  console.log("got load from TEAMS", data);
  logItem("OnLoad", "blue", "Started for " + data.entityId);
  let newItem = logItem("OnLoad", "blue", "Completed for " + data.entityId);
  setItems((oldItems) => [...oldItems, newItem]);
  microsoftTeams.appInitialization.notifySuccess();
};

export const Tab = () => {
  const [items, setItems] = useState<string[]>([]);
  const [title, setTitle] = useState("initial title");
  const [initState, setInitState] = useState(false);
  React.useEffect(() => {
    if (!initState) {
      return;
    }
    // get context
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      if (context) {
        console.log("got context from TEAMS", context);
        const newItem = logItem("Success", "green", "Loaded Teams context");
        setItems((oldItems) => [...oldItems, newItem]);
        setTitle(context.entityId);
        const newItem2 = logItem("FrameContext", "orange", "Frame context is " + context.frameContext);
        setItems((oldItems) => [...oldItems, newItem2]);
        if (context.frameContext === "sidePanel") {
          // ############################################
          // OnBeforeUnload
          microsoftTeams.registerBeforeUnloadHandler((readyToUnload) => {
            const result = beforeUnloadHandler(setItems, readyToUnload);
            return result;
          });

          // ############################################
          // OnLoad
          microsoftTeams.registerOnLoadHandler((data) => {
            loadHandler(setItems, data);
          }); 
          const newItem = logItem("Handlers", "orange", "Registered load and before unload handlers. Ready for app caching.");
          setItems((oldItems) => [...oldItems, newItem]);         
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
      microsoftTeams.appInitialization.notifySuccess();
      setInitState(true);
    }, timeout);
  }, []);


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
