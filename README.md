# Outlook GPT Rephraser Add-in

This is an add-in for Microsoft Outlook that allows users to rephrase a selected text using the OpenAI language model.

## How to install

1. Open Outlook and go to "Home" tab.
2. Click on "Get Add-ins".
3. Open "My add-ins" and scroll down to "Custom Addins".
4. Click on "Add a custom add-in" and select "Add from URL..."
5. Insert manifest URL:
```bash
https://stefanoaldegheri.github.io/OutlookGPTRephraser/manifest.xml
```
6. Click OK and return to main form.
7. Click on "New Mail".
8. Select tab button "GPT Rephraser" to open taskpane.
9. Open "Options" tab.
10. Insert "OpenAI GPT Key" from your personal OpenAI account.
11. Click "Save Options".

## How to use 

1. Click on "New Mail".
2. Select tab button "GPT Rephraser" to open taskpane.
3. Select the text to be rephrased in the mail body.
4. Define format.
5. Click "Rephrase" and review result.
6. Click "Update Selection" to replace the text in the mail.

## Developer instructions

1. Follow instructions at [Build yor first Outlook add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=yeomangenerator)
2. Clone this repository
3. Execute in Powershell
```bash
 cd repository-folder
 npm install office-addin-debugging
 code .
```
4. To sideload the add-in, replace the content in the '**manifest.xml**' file with the content found in the '**manifest_localhost.xml**' file.
5. Update code in webpack.config.js
```bash
const urlDev = "https://localhost:3000/";
const urlProd = "https://localhost:3000/";
```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
