import { ChoiceGroup, IChoiceGroupOption, Icon, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { AppInsightsContext, AppInsightsErrorBoundary } from "@microsoft/applicationinsights-react-js";
import { FunctionListener, ILogEntry, Logger, LogLevel } from "@pnp/logging";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import { IClientsidePage } from "@pnp/sp/clientside-pages";
import * as strings from "PanelCommandSetStrings";
import * as React from "react";
import { reactPlugin } from "../../../utils/AppInsights";
import { handleError } from "../../../utils/ErrorHandler";
import StatefulPanel from "../StatefulPanel/StatefulPanel";
import { ICopyPageCompProps } from "./ICopyPageCompProps";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import { ISearchResult, SearchResults } from "@pnp/sp/search";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { PermissionKind } from "@pnp/sp/security";

interface SearchResult {
	key: string;
	text: string;
}

export default function CopyPageComp(props: ICopyPageCompProps): JSX.Element {
	const [postAsNews, setPostAsNews] = useBoolean(true);
	const [hideSpinner, setHideSpinner] = useBoolean(true);
	const [refreshPage, setRefreshPage] = useBoolean(false);
	const [formDisabled, setFormDisabled] = useBoolean(true);
	const [itemId, setItemID] = React.useState(null);
	const [destinationSite, setDestinationSite] = React.useState<string>("");
	const [statusTxt, setStatusTxt] = React.useState<string>("");
	const [pageUrl, setPageUrl] = React.useState<string>("");
	const [newPageUrl, setNewPageUrl] = React.useState<string>("");
	const [options, setOptions] = React.useState<IChoiceGroupOption[]>([]);
	const [pageName, setPageName] = React.useState<string>("");
	const [newPageName, setNewPageName] = React.useState<string>("");
	const [statusType, setStatusType] = React.useState<MessageBarType>(null);

	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	const funcListener = new (FunctionListener as any)((entry: ILogEntry) => {
		switch (entry.level) {
			case LogLevel.Error:
				setStatusTxt(entry.message);
				setStatusType(MessageBarType.error);
				break;
			case LogLevel.Warning:
				setStatusTxt(entry.message);
				setStatusType(MessageBarType.warning);
				break;
		}
	});
	const getSP = (siteUrl: string): SPFI => {
		const sp = spfi(siteUrl).using(SPFx(props.context));
		return sp;
	}

	const mapSearchResultsToKeyText = (json: SearchResults): SearchResult[] => {
		return json.PrimarySearchResults
			.filter((item: ISearchResult) => item.Path !== props.context.pageContext.site.absoluteUrl)
			.sort((a: ISearchResult, b: ISearchResult) => a.Path.localeCompare(b.Path))
			.map((item: ISearchResult) => {
				return {
					key: item.Path,
					text: `${item.Title} (${item.Path})`
				};
			});
	}

	const getSites = async (): Promise<void> => {
		setHideSpinner.setFalse();
		const globalSite = getSP("https://borregaard.sharepoint.com");

		const globalDepartmentId = "d6faf00a-f32b-4d6d-b3d0-18c381d6f3d3";
		const noDepartmentId = "79d0ace8-f607-4c26-9c8b-77bc18796a32";
		const searchQuery = {
			Querytext: `(DepartmentId:${globalDepartmentId} OR DepartmentId: ${noDepartmentId}) AND contentclass:STS_Site`,
			RowLimit: 150,
			SelectProperties: ["Title", "Path", "DepartmentId"],
		};
		const results = await globalSite.search(searchQuery);
		console.log(results);
		const options = mapSearchResultsToKeyText(results);
		const filteredOptions = [];
		for (const option of options) {
			const siteUrl = option.key;
			const siteSP = getSP(siteUrl);

			try {
				const sitePagesLibrary = await siteSP.web.lists.ensureSitePagesLibrary();
				const currentUserHasWrite = await sitePagesLibrary.currentUserHasPermissions(PermissionKind.AddListItems);

				if (!currentUserHasWrite) {
					console.log(`User does not have write access to Site Pages on ${siteUrl}`);
				} else {
					console.log(`User has read access to Site Pages on ${siteUrl}`);
					filteredOptions.push(option);
				}
			} catch (error) {
				console.error(`Failed to check permissions for ${siteUrl}:`, error);
			}
		}


		setOptions(filteredOptions);
		setFormDisabled.setFalse();
		setHideSpinner.setTrue();
	}

	React.useEffect(() => {
		Logger.subscribe(funcListener);
		const fetchData = async (): Promise<void> => {
			try {
				await getSites();
			} catch (error) {
				console.error("Error fetching sites:", error);
			}
		};

		fetchData().catch((error) => console.error("Error in fetchData:", error));
		if (props.selectedRows.length === 1) {
			const selectedRow = props.selectedRows[0];
			setItemID(selectedRow.getValueByName("ID"));
		}

	}, []);

	const _errorFallback = (error: Error, info: { componentStack: string }): JSX.Element => {
		Logger.error(error);
		Logger.write(info.componentStack, LogLevel.Error);
		return (
			<MessageBar
				messageBarType={MessageBarType.error}
				isMultiline={true}
				dismissButtonAriaLabel="Close"
			>
				{error}
			</MessageBar>
		);
	};

	const SafeFilename = (fileName: string, convertToLower?: boolean): string => {
		const name = (fileName || "")
			.replace(/[^ a-zA-Z0-9-_]*/gi, "")
			.replace(/ /gi, "-")
			.replace(/-*$/gi, "")
			.replace(/[-]{2,}/gi, "-");
		return convertToLower ? name.toLowerCase() : name;
	}

	const copyPage = async (sourcePageUrl: string, targetSiteUrl: string): Promise<boolean> => {
		try {
			const page: IClientsidePage = await props.spfiContext.web.loadClientsidePage(sourcePageUrl);
			const newWeb = getSP(targetSiteUrl).web;
			const newPage = await page.copy(newWeb, SafeFilename(newPageUrl), newPageName, false);

			try {
				// if (newPage.sections.length > 1) {
				// 	newPage.sections[0].remove();
				// }
				newPage.commentsDisabled = page.commentsDisabled;
				newPage.layoutType = page.layoutType
				newPage.pageLayout = page.pageLayout;
				if (postAsNews) {
					await newPage.promoteToNews();
				}

				const pageItem = await (await newPage.getItem()).fieldValuesAsText();
				const pageUrl = pageItem.FileRef;
				await newPage.save(false);
				setPageUrl(pageUrl);
				setPageName(newPageName);
				setRefreshPage.setTrue();

			} catch (error) {
				console.error('Error copying page:', error);
			}
			return true;

		} catch (error) {
			console.error('Error copying page:', error);
			await handleError(error);
			return false;
		}
	}

	const onFormSubmitted = async (): Promise<void> => {
		setHideSpinner.setFalse();
		setFormDisabled.setTrue();
		const listItem = await props.spfiContext.web.lists.getById(props.listId).items.getById(itemId).select("FileRef")();
		const pageFileRef = listItem.FileRef;
		const sourcePageUrl = pageFileRef;

		const result: boolean = await copyPage(sourcePageUrl, destinationSite);
		if (props.onCompleted !== undefined) {
			props.onCompleted(result);
		}
		if (result) {
			setStatusTxt("The page has been copied successfully.");
			setStatusType(MessageBarType.success);

		} else {
			setStatusTxt("The page has not been copied successfully.");
			setStatusType(MessageBarType.error);
		}
		setHideSpinner.setTrue();
	};

	const onPanelDismissed = (): void => {
		if (refreshPage && props.panelConfig.onDismiss !== undefined) {
			props.panelConfig.onDismiss();
		}
	};
	const onChoiceChange = (event: React.FormEvent<HTMLInputElement>, option?: IChoiceGroupOption): void => {
		setDestinationSite(option?.key);
	};


	return (
		<AppInsightsContext.Provider value={reactPlugin}>
			<AppInsightsErrorBoundary
				onError={_errorFallback}
				appInsights={reactPlugin}
			>
				<StatefulPanel
					title={props.panelConfig.title}
					panelTop={props.panelConfig.panelTop}
					onDismiss={onPanelDismissed}
				>
					{statusTxt && (
						<MessageBar
							messageBarType={statusType}
							isMultiline={true}
							dismissButtonAriaLabel="x"
							onDismiss={() => setStatusTxt(null)}
						>
							{statusTxt}
						</MessageBar>
					)}
					<div style={{ padding: "20px" }}>Copy this page to: {destinationSite}</div>
					<div hidden={formDisabled}>
						<div style={{ padding: "10px" }}>
							<ChoiceGroup
								label="Which site?"
								options={options}
								onChange={onChoiceChange}
								required={true}
							/>
						</div>
						<div style={{ padding: "10px" }}>
							<TextField onChange={(event, newValue) => {
								setNewPageName(newValue || ''); setNewPageUrl(newValue || '');
							}} label="New page title:"
								required={true}></TextField>
						</div>
						<div style={{ padding: "10px" }} hidden={true}>
							<TextField onChange={(event, newValue) => {
								setNewPageUrl(newValue || '');
							}} label="New page url:" required={false} ></TextField>
						</div>
						<div style={{ padding: "10px" }}>
							<Toggle onChange={() => {
								setPostAsNews.toggle();
							}} label="Post as news:" checked={postAsNews} ></Toggle>
						</div>
						<PrimaryButton
							text={strings.btnSubmit}
							onClick={onFormSubmitted}
							allowDisabledFocus
							disabled={newPageUrl.length === 0 || destinationSite.length === 0}
							hidden={formDisabled}
						/></div>
					<div style={{ padding: "20px", fontSize: "18px" }} hidden={pageUrl.length === 0}><Icon iconName="PageLink"></Icon> Go to the copied page: <a href={pageUrl} target="_blank">{pageName}</a></div>
					<div style={{ padding: "20px" }} hidden={hideSpinner}> <Spinner size={SpinnerSize.large} /></div>
				</StatefulPanel>
			</AppInsightsErrorBoundary>
		</AppInsightsContext.Provider>
	);
}

