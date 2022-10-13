import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { BrowserRouter as Router, Route, Switch } from "react-router-dom";
import { Tab } from "./Tab";
import { TabConfig } from "./TabConfig";

function App() {
  microsoftTeams.initialize();
  return (
    <Router>
      <Switch>
        <Route path="/config">
          <TabConfig />
        </Route>
        <Route path="/tab">
          <Tab />
        </Route>
      </Switch>
    </Router>
  );
}

export default App;
