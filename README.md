# OutlookMD

OutlookMD is an Office Add-in for Microsoft Outlook that allows you to extract, format, and send meeting or message details in Markdown format. It is designed to work with both reading and composing messages and appointments.

## Features

- Extracts details from Outlook messages and appointments (subject, attendees, body, etc.)
- Formats meeting details as Markdown using Liquid templates
- Sends message or appointment data to a REST API endpoint
- Provides a custom command button in Outlook for easy access
- Works in both read and compose modes

## Getting Started

### Prerequisites

- Node.js (v16 or later recommended)
- Microsoft Outlook (desktop)
- Office Add-in sideloading enabled

### Installation

1. Clone this repository:
   ```sh
   git clone https://github.com/OfficeDev/Office-Addin-TaskPane.git
   cd Office-Addin-TaskPane
   ```
2. Install dependencies:
   ```sh
   npm install
   ```
3. Build the project:
   ```sh
   npm run build:dev
   ```
4. Start the add-in for Outlook:
   ```sh
   npm run start -- desktop --app outlook
   ```

### Sideloading the Add-in

- The add-in will be sideloaded into Outlook automatically when running the start command above.
- If prompted, allow loopback for Microsoft Edge WebView.

## Usage

- Open Outlook (desktop).
- Select or compose a message or appointment.
- Use the custom button added by the add-in to extract and send details.
- A notification will confirm when the action is complete.

## Project Structure

- `src/commands/` – Command logic and REST integration
- `src/taskpane/` – Task pane UI and supporting files
- `manifest.xml` – Office Add-in manifest
- `webpack.config.js`, `babel.config.json`, `tsconfig.json` – Build configuration

## Development

- Edit `src/commands/commands.ts` to customize extraction or API integration.
- Update `manifest.xml` to change add-in metadata or commands.
- Use `npm run lint` and `npm run lint:fix` to check and fix code style.

## Deployement

The addin is deployed to an Azure storage account after setting the production url in the manifest.

```
npm run build
az storage blob upload-batch --account-name storageoutlookmd --destination '$web' --overwrite --source ./dist/
```

## License
MIT License. See [LICENSE](LICENSE) for details.