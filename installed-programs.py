"""Detect installed programs and save to csv. Optionally only save programs not in an existing csv file.

Usage:
  installed-programs.py <output> [--compare-to=<input>]
  installed-programs.py (-h | --help)

Arguments:
  <output>  Output file
  --compare-to=<input>  Output only entries that are missing from the input file

Options:
  -h, --help     Show this screen.

"""

import winreg
import csv
from docopt import docopt


def get_installed(hive, flag):
    aReg = winreg.ConnectRegistry(None, hive)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                          0, winreg.KEY_READ | flag)

    count_subkey = winreg.QueryInfoKey(aKey)[0]

    software_list = []

    for i in range(count_subkey):
        software = {}
        try:
            asubkey_name = winreg.EnumKey(aKey, i)
            asubkey = winreg.OpenKey(aKey, asubkey_name)
            software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]

            try:
                software['publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]
            except EnvironmentError:
                software['publisher'] = 'undefined'
            software_list.append(software)
        except EnvironmentError:
            continue

    return software_list


def save_to_csv(software_list, in_file):
    with open(in_file, "w", newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Name', 'Publisher'])
        for software in software_list:
            writer.writerow([software['name'], software['publisher']])


def read_from_csv(out_file):
    with open(out_file, "r", newline='') as file:
        load_soft = []
        reader = csv.reader(file)
        next(reader, None)
        for row in reader:
            load_soft.append({
                "name": row[0],
                "publisher": row[1]
            })
        return load_soft


if __name__ == "__main__":
    arguments = docopt(__doc__)
    if arguments["--compare-to"] is None:
        installed_software = get_installed(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + get_installed(
            winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY) + get_installed(winreg.HKEY_CURRENT_USER, 0)

        save_to_csv(installed_software, arguments["<output>"])
        print(f"Total apps detected: {len(installed_software)}")

    else:
        missing_list = []
        installed_software = get_installed(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + get_installed(
            winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY) + get_installed(winreg.HKEY_CURRENT_USER, 0)
        read_file = read_from_csv(arguments["--compare-to"])

        for software in installed_software:
            if software not in read_file:
                missing_list.append(software)

        print(f"Total missing from input detected: {len(missing_list)}")
        if missing_list:
            save_to_csv(missing_list, arguments["<output>"])
