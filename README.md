---
page_type: sample
urlFragment: excel-blazor-add-in + ribbon derived from Maarten van Stam sample
products:
  - office-add-ins
  - office
languages:
  - typescript
  - C#
extensions:
  contentType: samples
  technologies: Add-ins
  createdDate: '03/21/2023 10:00:00 PM'
description: 'Create a Blazor Webassembly Excel add-in showcasing some samples + ribbon.'
---

# Create a Blazor Webassembly Excel add-in + ribbon

This sample shows how to build an Excel add-in using .NET Blazor technologies with typescript. Blazor Webassembly allows you to build Office Add-ins using .NET, C#, and typescript to interact with the Office JS API. The add-in uses typescript to work with the document and Office JS APIs, but you build the user interface and all other non-Office interactions in C# and .NET Core Blazor technologies.
 
- Work with Blazor Webassembly in the context of Office.
- Build cross-platform Office Add-ins using Blazor, C# and typescript Interop.
- Initialize the Office typescript API library in Blazor context.
- Interact with Excel to manipulate worksheets.
- Interact with workbook content through Office typescript APIs.
+ a demonstration of office ribbon

## Applies to

- Excel on the web, Windows, and Mac.

## Prerequisites

- Microsoft 365 - Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone the [Office Add-ins samples repository](https://github.com/OfficeDev/Office-Add-in-samples).
1. Open Visual Studio 2022 and open the **Office-Add-in-samples\Samples\blazor-add-in\excel-blazor-add-in\excel-blazor-add-in.sln** solution. (Do **not** open Visual Studio "as administrator". There is a bug that will prevent the add-in from sideloading when Visual Studio is elevated in this way.)
1. Choose **Debug** > **Start Debugging**. Or press F5 to start the solution.
1. When Excel opens, choose **Home** > **Show Taskpane**.

CAUTION : if you change anything in the XML you may have to restart your machine 

Next, try out the controls.

## Understand an Office Add-in in Blazor Context

An Office Add-in is a web application that extends Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application, you must provide a web server to host the files.
Building the Office Add-in as a Blazor Webassembly allows you to build a .NET Core compliant website that interacts with the Office JS APIs. If your background is with VBA, VSTO, or COM add-in development, you may find that building Office Add-ins using Blazor Webassembly is a familiar development technique.

## Key parts of this sample

This sample uses a Blazor Webassembly file that runs cross-platform in various browsers supporting WASM (Webassembly). The Blazor WASM App demonstrates some basic Excel functions.

The purpose of this sample is to show you how to build and interact with the Blazor, C# and typescript Interop options. If you're looking for more examples of interacting with Excel and Office JS APIs, see [Script Lab](https://aka.ms/getscriptlab).


### Blazor pages

The **Pages** folder contains the Blazor pages, such as **Index.razor**. Each **.razor** page also contains two code-behind pages, for example, **Index.razor.cs** and **Index.razor.js**. The C# file first establishes an interop connection with the typescript file.


```
In BlazorAddin lib.module.js the code is modified in order to postpone blazor launch
 
 export async function afterStarted(blazor) {
    var customScript = document.createElement('script');
    customScript.setAttribute('src', 'Pages/Index.razor.js');
    document.head.appendChild(customScript);
    console.log("We are now entering function: afterStarted");
}

```

The typescript catch the ribbon event to interact with the C# and returns.

```typescript
(window: any).lancerchaîné = async (event) => {
    console.log("lancerchaîné");
    try {
        callStaticLocalComponentMethod();
    }
    catch (err) {
        console.log();
        console.log("erreur lancer : " + err.message);
    }
    console.log("eventlancer Completed");
    event.completed();
}
For any events that need to interact with the Office document, the typescript file calls through interop to the C# file.

```typescript/typescript
  await DotNet.invokeMethodAsync("BlazorAddIn", "Localfunction");

 
```

The fundamental pattern includes the following steps.



## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

Solution | Authors
---------|----------
Create a Blazor Webassembly Excel add-in | [Maarten van Stam](https://mvp.microsoft.com/en-us/PublicProfile/33535) + RaoulRSV

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | March 21, 2023 | Derived from Maarten van Stam sample

## Copyright

Copyright(c) Maarten van Stam. All rights reserved.Licensed under the MIT License.+ RaoulRSV for office ribbon

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

