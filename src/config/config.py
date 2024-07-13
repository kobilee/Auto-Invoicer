import json
import os

def load_settings(json_path):
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f"The configuration file '{json_path}' was not found.")
    
    with open(json_path, 'r') as file:
        return json.load(file)


def save_settings(config, json_path):
    with open(json_path, 'w') as file:
        json.dump(config, file, indent=4)
