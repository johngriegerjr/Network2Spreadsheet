# Network2Spreadsheet
This script connects to a list of network devices, runs a show command, transforms the output of the show command using TextFSM, then creates a workbook with the output from the devices in a sheet per device type.
## Instructions
1. Clone repo
     * git clone https://github.com/johngriegerjr/Network2Spreadsheet.git
2. Create Virtual Environment
     * **Mac**:
       * cd Network2Spreadsheet
       * python3 -m venv .
       * source ./bin/activate
       * pip install --upgrade pip
     * **Windows**
       * cd Network2Spreadsheet
       * python -m venv my_env
       * .\Scripts\activate
       * pip install --upgrade pip
3. Install requirements.txt
     * pip install -r requirements.txt
4. Edit device_list_both.csv
     * **device_name** is a friendly name that will appear in the outputed workbook
     * **device_type** defines the Netmiko device driver to be used. Each unique device_type will have its own spreadsheet in the outputted workbook.
     * **host** can be a DNS name or IP address that will be used to SSH to the network device.
5. Execute the script
     * Run script
       * python Network2Spreadsheet.py
     * Enter username
     * Enter password
       * This password will also be used for enable
     * Enter show command to run
       * This needs to be a command that has a TextFSM template created for
6. View results in output folder
     * cd output
     * ls
