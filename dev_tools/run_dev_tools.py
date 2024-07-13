import json
import fitz
from dev_tools import DevTools

def load_dev_config(json_path):
    with open(json_path, 'r') as file:
        return json.load(file)

def run_find_direction_tool(dev_tools, config):
    file_path = config["file_path"]
    searches = config["searches"]

    with fitz.open(file_path) as pdf_file:
        for search in searches:
            string = search["string"]
            direction = search["direction"]
            page_num = search["page"]

            page = pdf_file[page_num]
            search_positions = page.search_for(string)

            if not search_positions:
                print(f"String '{string}' not found on page {page_num}")
                continue

            print(f"\nSearch results for '{string}' on page {page_num} in direction '{direction}':")
            for pos in search_positions:
                found_text, rect, offset = dev_tools.find_direction(page, pos, direction)
                if found_text:
                    print(f"    Found Text: {found_text}")
                    print(f"    Offset: {offset}")
                    print(f"    Rect: {rect}\n")
                else:
                    print(f"  Text not found within {dev_tools.max_search_distance} units {direction} of {pos}")

def run_dev_tools(config_path):
    config = load_dev_config(config_path)
    tool = config["tool"]

    dev_tools = DevTools()

    if tool == "find_direction":
        run_find_direction_tool(dev_tools, config)
    else:
        print(f"Unknown tool: {tool}")

if __name__ == "__main__":
    config_path = "dev_config.json"
    run_dev_tools(config_path)
