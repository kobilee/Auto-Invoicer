import json
import os

def load_settings(json_path, document_type=None):
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f"The configuration file '{json_path}' was not found.")
    
    with open(json_path, 'r') as file:
        config = json.load(file)
    
    base_dir = os.getcwd()

    if document_type == 'invoice':
        config['input'] = os.path.join(base_dir, config['input_invoices'])
    elif document_type == 'statement':
        config['input'] = os.path.join(base_dir, config['input_statements'])
    
    config['excel'] = os.path.join(base_dir, config['database_dir'], config['database_filename'])

    if not os.path.exists(config['backup']):
        os.makedirs(config['backup'])

    return config


def save_settings(config, json_path):
    with open(json_path, 'w') as file:
        json.dump(config, file, indent=4)
