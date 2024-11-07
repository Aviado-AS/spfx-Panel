import { AnalyticsPlugin } from "@microsoft/applicationinsights-analytics-js";
import { override } from "@microsoft/decorators";
import { BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters, ListViewStateChangedEventArgs } from "@microsoft/sp-listview-extensibility";
import { ConsoleListener, Logger } from "@pnp/logging";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import * as strings from "PanelCommandSetStrings";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AppInsights } from "../utils/AppInsights";
import { AppInsightsLogListener } from "../utils/AppInsightsLogListener";
import { ICopyPageCompProps } from "./components/CopyPageComp/ICopyPageCompProps";
import CopyPageComp from "./components/CopyPageComp/CopyPageComp";

export interface IPanelCommandSetProperties {
	logLevel?: number;
	listName: string;
	appInsightsConnString?: string;
}
interface IProcessConfigResult {
	visible: boolean;
	disabled: boolean;
	title: string;
}

const LOG_SOURCE: string = "PanelCommandSet";

export default class PanelCommandSet extends BaseListViewCommandSet<IPanelCommandSetProperties> {
	//#region variables
	private panelPlaceHolder: HTMLDivElement = null;
	private panelTop: number;
	private panelId: string;
	private compId: string;
	private spfiContext: SPFI;
	private appInsights: AnalyticsPlugin;
	//#endregion

	@override
	public onInit(): Promise<void> {
		//list name registered in Extension properties
		const _isListRegistered = this.context.listView.list.serverRelativeUrl.indexOf("/SitePages") > -1 ? true : false;

		const _setLogger = (appInsights?: AnalyticsPlugin): void => {
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			Logger.subscribe(new (ConsoleListener as any)());
			if (appInsights !== undefined) {
				Logger.subscribe(new AppInsightsLogListener(appInsights));
			}

			if (this.properties.logLevel && this.properties.logLevel in [0, 1, 2, 3, 99]) {
				Logger.activeLogLevel = this.properties.logLevel;
			}

			Logger.write(`${LOG_SOURCE} Activated Initialized with properties:`);
			Logger.write(`${LOG_SOURCE} ${JSON.stringify(this.properties, undefined, 2)}`);
		};
		const _setPanel = (): void => {
			this.panelTop = document.querySelector("#SuiteNavWrapper").clientHeight;
			this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
		};
		const _setCommands = (): void => {
			const _setCommandState = (command: Command, config: IProcessConfigResult): void => {
				command.visible = config.visible;
				command.disabled = config.disabled;
				command.title = config.title;
			};

			const compareTwoCommand: Command = this.tryGetCommand("COMMAND_1");
			if (compareTwoCommand) {
				_setCommandState(compareTwoCommand, {
					title: "Copy page to another site",
					visible: true,
					disabled: true,
				});
			}

			this.raiseOnChange();
		};

		if (!_isListRegistered) {
			return;
		}

		if (this.properties.appInsightsConnString) {
			this.appInsights = AppInsights(this.properties.appInsightsConnString);
			// this.appInsights.trackPageView(); //DON'T, this will record page view twice
			_setLogger(this.appInsights);
		} else {
			_setLogger();
		}

		_setPanel();
		_setCommands();
		this.spfiContext = spfi().using(SPFx(this.context));
		this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

		return Promise.resolve();
	}

	// Triggered when row(s) un/selected
	public _onListViewStateChanged(args: ListViewStateChangedEventArgs): void {
		const itemSelected = this.context.listView.selectedRows && this.context.listView.selectedRows.length === 1;

		const compareTwoCommand: Command = this.tryGetCommand("COMMAND_1");
		if (compareTwoCommand && compareTwoCommand.disabled === itemSelected) {
			compareTwoCommand.disabled = !itemSelected;
			this.raiseOnChange();
		}

		//#region NOTE
		// NOTE: use raiseOnChange() with caution; frequent calls can lead to low performance of the list
		// https://github.com/SharePoint/sp-dev-docs/discussions/7375#discussioncomment-2053604
		//#endregion
	}

	@override
	public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
		const _showComponent = (props: ICopyPageCompProps): void => {
			//#region comments
			this.compId = Date.now().toString();
			this.panelPlaceHolder.setAttribute("id", this.compId);
			const element: React.ReactElement<ICopyPageCompProps> = React.createElement(CopyPageComp, {
				...props,
				key: this.compId,
			});

			ReactDOM.render(element, this.panelPlaceHolder);
		};



		const _dismissPanel = (): void => {
			Logger.write(strings.lblRefreshing);
			location.reload();
		};
		const _onCompleted = (success: boolean): void => {
			if (this.appInsights !== undefined) {
				this.appInsights.trackEvent({
					name: success ? strings.lblItemUpdate_OK : strings.lblItemUpdate_Err,
				});
			}
		};

		switch (event.itemId) {
			case "COMMAND_1":
				_showComponent({
					panelConfig: {
						panelTop: this.panelTop,
						title: "COPY THIS PAGE TO ANOTHER SITE",
						onDismiss: _dismissPanel,
					},
					spfiContext: this.spfiContext,
					listName: this.context.listView.list.title,
					listId: this.context.listView.list.guid.toString(),
					selectedRows: event.selectedRows,
					context: this.context,
					onCompleted: _onCompleted,
				});
				break;
			default:
				throw new Error("Unknown command");
		}
	}

	public onDispose(): void {
		ReactDOM.unmountComponentAtNode(document.getElementById(this.panelId));
		ReactDOM.unmountComponentAtNode(document.getElementById(this.componentId));
	}
}
