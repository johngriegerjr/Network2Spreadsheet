# Author:     John Grieger
# Version:    1.0
# Date:       20250913
# File:       Netwrk2Spreadsheet.py
# Copyright:  Copyright (c) 2025 John Grieger

"""
This script uses threads to pull show command information from a list of devices, transforms the output,
and writes the information to an Excel file
"""

import csv
import getpass
import os
from pandas import ExcelWriter as pdxl
from pandas import DataFrame as pddf
from datetime import datetime
from netmiko import ConnectHandler
from pathlib import Path
from netmiko import ConnectHandler, NetMikoAuthenticationException
from concurrent.futures import ThreadPoolExecutor, as_completed, wait

def get_show_command(netmiko_dict, show_command, device_name):
    try:
        # Create SSH session to device
        ssh_connection = ConnectHandler(**netmiko_dict)
        # Escalate to privileged mode
        ssh_connection.enable()
        result_list = [{"dev_hostname": device_name, "device_type": netmiko_dict['device_type']}]
        # Send show command with increased delay to ensure all results are captured and use textfsm
        output = ssh_connection.send_command(show_command, use_textfsm=True, delay_factor=2, read_timeout=900)
        # Append output to results list
        result_list.append(output)
        # Disconnect SSH connection
        ssh_connection.disconnect()
        return result_list

    except NetMikoAuthenticationException as ex:
        print(f"{netmiko_dict['host']}\tCredentials failed with\t{netmiko_dict['username']}")
        result_list = [{"dev_hostname": device_name, "device_type": netmiko_dict['device_type']}]
        return result_list
        pass

    except Exception as ex:
        print(f"Something went wrong connecting to\t{netmiko_dict['host']}\t{ex}")
        result_list = [{"dev_hostname": device_name, "device_type": netmiko_dict['device_type']}]
        return result_list
        pass

def create_netmiko_dict(csv_list,un,pw):
    netmiko_list = []
    with open(csv_list) as f:
        read_csv = csv.DictReader(f)
        # Iterate through list of devices
        for entry in read_csv:
            netmiko_list.append({entry['device_name']: {'host': entry['host'], 'device_type': entry['device_type'], 'username': un, 'password': pw, 'secret': pw}})
    return netmiko_list

def main():
    # Pompt user for data
    un = input('Enter Username: ')
    pw = getpass.getpass(prompt='Enter Password: ')
    sh_cmd = input('Enter show command to run: ')
    # Specify CSV file with list of devices to connect to
    device_list = "device_list_both.csv"
    # Create name for output file
    file_name = datetime.now().strftime(f'{sh_cmd}_output-%Y-%m-%d_%H-%M.xlsx')
    Path('output').mkdir(parents=True, exist_ok=True)
    log_dir = os.path.join(os.getcwd(), 'output')
    log_fname = os.path.join(log_dir, file_name)

    netmiko_list = create_netmiko_dict(device_list,un,pw)

    try:
        processes = []
        with ThreadPoolExecutor(max_workers=10) as executor:
            for device in netmiko_list:
                for device_name,netmiko_dict in device.items():
                    print(device_name)
                    processes.append(executor.submit(get_show_command, netmiko_dict, sh_cmd, device_name))

            wait(processes)
    # Allow all exceptions
    except Exception as ex:
        print(f"Exception occurred: {ex}\r")
        pass

    g = globals()
    device_type_set = set()
    for task in as_completed(processes):
        device_type_set.add(task.result()[0]['device_type'])

    for i in device_type_set:
        g["new_" + i + "_list"] = []
    for task in as_completed(processes):
        if task.result():
            for dev_type in device_type_set:
                if dev_type in task.result()[0]['device_type']:
                    check_len = len(task.result())
                    if check_len > 1:
                        for dict in task.result()[1]:
                            new_result = {"dev_hostname": task.result()[0]['dev_hostname']}
                            new_result.update(dict)
                            g['new_'+dev_type+'_list'].append(new_result)
                    else:
                        for dict in task.result()[0]:
                            new_result = {"dev_hostname": task.result()[0]['dev_hostname']}
                            #new_result.update(dict)
                            g['new_'+dev_type+'_list'].append(new_result)

    # Write Pandas data frames to Excel workbook
    with pdxl(log_fname) as writer:
        for i in device_type_set:
            dframe = pddf.from_records(g['new_'+i+'_list'])
            dframe.to_excel(writer, index=False, sheet_name=f'{i}')

if __name__ == "__main__":
    main()
