import "azure-devops-ui/Core/override.css";
// import "azure-devops-ui/Core/_platformCommon.scss"
import "es6-promise/auto";
import type * as React from "react";
import { createRoot } from "react-dom/client";
import * as ReactDOM from "react-dom";
import "./Common.scss";
// import * as style from "azure-devops-ui/Core/_platformCommon.scss";

export function showRootComponent(component: React.ReactElement<unknown>) {
	const root = document.getElementById("root");
	if (!root) throw new Error("root element not found");
	createRoot(root).render(component);
}
