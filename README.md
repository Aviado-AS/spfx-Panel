# spfx-panel

## Summary

This control renders stateful Panel that can be used with [ListView Command Set extensions](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api). It may optionally refresh the list view page after the panel is closed.
It opens when a List Command button is clicked, and closes using either Panel's close button, or on "light dismiss".

It may be used to replace Dialog component, ensuring the User Interface is consistent with that of SharePoint Online.


[picture of the solution in action, if possible]


## Compatibility

![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
[![version](https://img.shields.io/badge/SPFx-1.13.1-green)](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)  ![version](https://img.shields.io/badge/Node.js-14.15.0-green)
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Local Workbench Incompatible](https://img.shields.io/badge/Local%20Workbench-Incompatible-red.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

SPFx 1.13 does not support local workbench. To test this solution you must have a SharePoint site.

## Solution

Solution|Author(s)
--------|---------
folder name | Kinga Kazala

## Version history

Version|Date|Comments
-------|----|--------
1.0|January 29, 2021|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **nvm use 14.15.0**
  - **npm install**
  - **gulp serve --nobrowser**
  - debug

## Features

Opening and closing Panel is a no-brainer as long as it is controled by a higher component.
In the case of ListView Command Set, controlling Panel state requires slightly more effort.

This extension illustrates the following concepts:

- Panel component with (optionally, recommended) Error Boundary
- Logging using  @pnp/logging Logger

### React Error Boundary

As of React 16, it is recommended to use error boundaries for handling errors in the component tree.
Error boundaries **do not catch** errors for event handlers, asynchronous code, server side rendering and errors thrown in the error boundary itself; try/catch is still required in these cases.
This solution uses [react-error-boundary](https://www.npmjs.com/package/react-error-boundary) component.

### PnP Logger

Logging is implemented using [@pnp/logging](https://pnp.github.io/pnpjs/logging) module. [Log level](https://pnp.github.io/pnpjs/logging/#log-levels) is defined as a customizer property, which allows changing log level of productively deployed solution, in case troubleshooting is required.

Errors returned by [@pnp/sp](https://pnp.github.io/pnpjs/sp/#pnpsp) commands are handled using `Logger.error(e)`, which parses and logs the error message. If the error message should be displayed in the UI, use the [handleError](src\common\errorhandler.ts) function  implemented based on [Reading the Response](https://pnp.github.io/pnpjs/concepts/error-handling/#reading-the-response) example.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [SPFx Debug Configuration](https://marketplace.visualstudio.com/items?itemName=eliostruyf.spfx-debug) - Visual Studio Code extension to add the required configuration for debugging SPFx solutions
- [PnP/PnPjs Getting Started](https://pnp.github.io/pnpjs/getting-started/)
- [PnP/PnPjs Error Handling](https://pnp.github.io/pnpjs/concepts/error-handling/)
- [Error Boundaries](https://reactjs.org/docs/error-boundaries.html) in React 16, and [react-error-boundary](https://www.npmjs.com/package/react-error-boundary) component
- [I Made a Tool to Generate Images Using Office UI Fabric Icons](https://joshmccarty.com/made-tool-generate-images-using-office-ui-fabric-icons/) to generate CommandSet icons