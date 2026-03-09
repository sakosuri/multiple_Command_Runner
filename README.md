multiple_Command_Runner
Application which can execute multiple commands on all the reachable devices on the Cisco Catalyst Center

This script will perform the below actions:
1. Authenticate with Catalyst Center and generates a Token to execute all the Catalyst Center API
2. Read all the Commands from the Excel file
3. Retrieves all the devices from the Catalyst Center and will check the no. of reachable devices
4. Verifies the supported commands from the commands list
5. Executes the supported commands on all the reachable devices
6. The output of all the commands will be saved in an excel sheet. These are the attributes of the excel sheet: Hostname, Device Id, Command, Command execution Status, Command Output

This script is a modified version of https://github.com/cisco-en-programmability/dnacenter_command_runner/tree/master which is created by Gabriel Zapodeanu. 
