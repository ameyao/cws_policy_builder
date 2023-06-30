import json

import requests
import yaml
from convert_excel_to_yaml import _convert_excel_to_yaml
from progress_bar import *

# COLOUR PALLET
HEADER = '\033[95m'
OKBLUE = '\033[94m'
OKCYAN = '\033[96m'
OKGREEN = '\033[92m'
WARNING = '\033[93m'
FAIL = '\033[91m'
ENDC = '\033[0m'
BOLD = '\033[1m'
UNDERLINE = '\033[4m'

with open(r'initial_setup.yaml') as file:
    yaml_data_init = yaml.load(file, Loader=yaml.FullLoader)

token = yaml_data_init['api_token']
vco_url = yaml_data_init['vco_url']
api_path = yaml_data_init['security_policy_api_path']
enterpriseLogicalId = yaml_data_init['enterpriseLogicalId']
excel_file = yaml_data_init['excel_file_path']
yaml_policy_file_path = yaml_data_init['yaml_file_path']


def convert_excel_to_yaml():
    converted_data = _convert_excel_to_yaml(excel_file)
    yaml_file_name = converted_data['security_policy']['payload0']['name'] + ".yaml"
    yaml_policy_file = yaml_policy_file_path + yaml_file_name
    with open(yaml_policy_file, "w") as policy_file:
        yaml.dump(converted_data, policy_file, default_flow_style=False)
        print(OKGREEN + "\tExcel file is converted to YAML\n\tFile name: {}".format(yaml_file_name) + ENDC)

    with open(yaml_policy_file, 'r') as policy_file:
        yaml_converted_data = yaml.load(policy_file, Loader=yaml.FullLoader)
        return yaml_converted_data


def vco_connectivity():
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies'

    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        print(OKGREEN + "\tVCO Connection successful" + ENDC)

    else:
        print(response.status_code)
        print(WARNING + "\tUnable to reach VCO" + ENDC)


def create_new_security_policy(yaml_converted_data):
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies'
    payload = yaml_converted_data['security_policy']['payload0']
    response = requests.post(
        api_url,
        headers=headers,
        json=payload
    )
    if response.status_code == 201:
        response = json.loads(response.text)
        print(OKGREEN + "\tCreated a new security policy" + ENDC)
        yaml_converted_data['security_policy']['policy_id'] = response['id']

    else:
        print(response.status_code)
        print(WARNING + "\tUnable to create a new security policy" + ENDC)


def create_url_filtering_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/urlFilteringRules'
    num_payloads = (len(yaml_converted_data['url_filtering_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['url_filtering_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue

        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new URL filtering policy" + ENDC)
    print(OKGREEN + "\tCreated {} URL filtering rules".format(num_payloads) + ENDC)


def create_content_filtering_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/contentFilteringRules'
    num_payloads = (len(yaml_converted_data['content_filtering_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['content_filtering_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue

        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new content filtering policy" + ENDC)
    print(OKGREEN + "\tCreated {} content filtering rules".format(num_payloads) + ENDC)


def create_geo_based_filtering_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/geoFilteringrules'
    num_payloads = (len(yaml_converted_data['geo_filtering_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['geo_filtering_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue

        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new GEO-Based filtering policy" + ENDC)
    print(OKGREEN + "\tCreated {} GEO-Based filtering rules".format(num_payloads) + ENDC)


def create_content_inspection_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/contentInspectionRules'
    num_payloads = (len(yaml_converted_data['content_inspection_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['content_inspection_policy']['payload{}'.format(payload_number)]

        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue
        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new Content Inspection rules" + ENDC)

    print(OKGREEN + "\tCreated {} Content Inspection rules".format(num_payloads) + ENDC)


def create_ssl_inspection_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/sslInspectionRules'
    num_payloads = (len(yaml_converted_data['ssl_inspection_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['ssl_inspection_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue
        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new SSL Inspection rules" + ENDC)

    print(OKGREEN + "\tCreated {} SSL Inspection rules".format(num_payloads) + ENDC)


def create_casb_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/CasbRules'
    num_payloads = (len(yaml_converted_data['casb_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['casb_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue
        else:
            print(response.status_code)
            print(WARNING + "\tUnable to create a new CASB rules" + ENDC)

    print(OKGREEN + "\tCreated {} CASB rules".format(num_payloads) + ENDC)


def create_dlp_rules(yaml_converted_data):
    cws_policy_id = yaml_converted_data['security_policy']['policy_id']
    headers = {'Authorization': f'Token {token}'}
    api_url = vco_url + api_path + enterpriseLogicalId + '/cwsPolicies/' + cws_policy_id + '/DlpRules'
    num_payloads = (len(yaml_converted_data['dlp_policy']))
    for payload_number in range(num_payloads):
        payload = yaml_converted_data['dlp_policy']['payload{}'.format(payload_number)]
        response = requests.post(
            api_url,
            headers=headers,
            json=payload
        )
        if response.status_code == 201:
            continue
        else:
            print(response.json())
            print(WARNING + "\tUnable to create a new DLP rules" + ENDC)

    print(OKGREEN + "\tCreated {} DLP rules".format(num_payloads) + ENDC)