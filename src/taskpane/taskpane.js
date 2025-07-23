/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import React from "react";
import ReactDOM from "react-dom";
import App from "./App.jsx";
import { createRoot } from "react-dom/client";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    const root = createRoot(document.getElementById("root"));
    root.render(<App />);
  }
});
