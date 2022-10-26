# PowerAutomateRunner  
A tool used to automate power automate desktop flows.

## Installation  
Download the [latest release](./releases/latest/download/PowerAutomateRunner.zip) and extract the zip archive to a folder of your choice.  

## Usage
This app runs using command line arguments.
| Argument | Description | Example | Required |
| --- | --- | --- | --- |
| `--flow-name` | The name of the flow to run. | `--flow-name "Test"` | Yes |
| `--retry-attempts` | The number of attempts to retry starting the flow if this tool fails to start the flow. | `--retry-attempts 10` | No (Default: `10`) |
| `--retry-interval` | The interval in milliseconds to wait between retry attempts. | `--retry-interval 250` | No (Default: `250`) |
| `-pause-on-error` | If this flag is set, the app will pause before exiting if an error occurs. | `-pause-on-error` | No (Default: `false`) |
