# SASE CWS Policy Builder

import requests
import urllib3
from progress_bar import wait
from cws_config_builder_functions import *

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
requests.packages.urllib3.disable_warnings()


# COLOR PALETTE
HEADER = "\033[95m"
OKBLUE = "\033[94m"
OKCYAN = "\033[96m"
OKGREEN = "\033[92m"
WARNING = "\033[93m"
FAIL = "\033[91m"
ENDC = "\033[0m"
BOLD = "\033[1m"
UNDERLINE = "\033[4m"


def cws_policy_builder():
    print(OKBLUE + "\n\t======== Step 1 - Convert CWS Policy Excel File into YAML ========" + ENDC)
    wait(2)
    yaml_converted_data = convert_excel_to_yaml()
    print(OKBLUE + "\n\t======== Step 2 - Test VCO connectivity ========" + ENDC)
    wait(2)
    vco_connectivity()
    print(OKBLUE + "\n\t======== Step 3 - Create new security policy ========" + ENDC)
    wait(2)
    create_new_security_policy(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 4 - Create URL Filtering Rules ========" + ENDC)
    wait(2)
    create_url_filtering_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 5 - Create Content Filtering Rules ========" + ENDC)
    wait(2)
    create_content_filtering_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 6 - Create GEO-Based Filtering Rules ========" + ENDC)
    wait(2)
    create_geo_based_filtering_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 7 - Create Content Inspection Rules ========" + ENDC)
    wait(2)
    create_content_inspection_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 8 - Create SSL Inspection Rules ========" + ENDC)
    wait(2)
    create_ssl_inspection_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 9 - Create CASB Rules ========" + ENDC)
    wait(2)
    create_casb_rules(yaml_converted_data)
    print(OKBLUE + "\n\t======== Step 10 - Create DLP Rules ========" + ENDC)
    wait(2)
    create_dlp_rules(yaml_converted_data)
    print(HEADER + "\n\t======== Completed ========" + ENDC)


if __name__ == "__main__":
    print(HEADER + "\n\t======== CWS Policy Builder ========" + ENDC)
    cws_policy_builder()
