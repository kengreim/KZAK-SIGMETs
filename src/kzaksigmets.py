from pathlib import Path
import os
import winreg
import requests
from lxml import etree
import argparse
from colorama import Fore, Back, Style
from decimal import Decimal
import traceback
import subprocess

# Constants
VATSYS_MAPS_PATH_RELATIVE = r'vatSys Files\Profiles\ATOP Oakland\Maps'
ISIGMET_API_URL = 'https://www.aviationweather.gov/cgi-bin/json/IsigmetJSON.php'
DEFAULT_FILENAME = 'SIGMET.XML'
DEFAULT_MAP_ATTRIBUTES = {
    'Type'             : 'Filled',
    'Name'             : 'SIGMETS',
    'Priorty'          : '1',
    'CustomColourName' : 'Indigo'
}

def error(error_message: str):
    print(Fore.WHITE + Back.RED + 'ERROR:' + Style.RESET_ALL + ' ' + error_message)

def log(log_message: str):
    print(Fore.WHITE + Back.GREEN + 'LOG:' + Style.RESET_ALL + ' ' + log_message)

def exit_with_wait():
    input('Press enter key to exit...')
    exit()

def find_vatsys_maps_dir() -> str | None:

    # First we will try the registry method
    try:
        home_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, home_path, access=winreg.KEY_READ)
        key_val, _ = winreg.QueryValueEx(key, 'Personal')
        winreg.CloseKey(key)
        full_dir = Path(key_val, VATSYS_MAPS_PATH_RELATIVE)
        if os.path.exists(full_dir):
            return full_dir
    except:
        pass

    # If we failed to find the folder via registry, try with Path method
    try:
        full_dir = Path(str(Path.home()), 'Documents', VATSYS_MAPS_PATH_RELATIVE)
        if os.path.exists(full_dir):
            return full_dir
    except:
        pass
    
    # If we are here, we haven't found anything, so return None
    return None

def find_vatsys_exec() -> str | None:

    # Try the x86 folder first
    full_path = Path(os.environ['ProgramFiles(x86)'], 'vatSys', 'bin', 'vatSys.exe')
    if os.path.exists(full_path):
        return full_path

    # Next try the regular Program Files folder
    full_path = Path(os.environ['ProgramW6432'], 'vatSys', 'bin', 'vatSys.exe')
    if os.path.exists(full_path):
        return full_path
    
    # Return none if both fail
    return None

def filter_kzak_sigmets(geojson_dict: dict) -> list[dict]:
    return_list = []
    for feature in geojson_dict['features']:
        if 'firId' in feature['properties'] and feature['properties']['firId'] == 'KZAK':
            return_list.append(feature)

    return return_list

def make_base_map_xml(map_attributes: dict[str, str] = DEFAULT_MAP_ATTRIBUTES) -> tuple[etree.Element, etree.Element]:
    maps_root = etree.Element('Maps')
    map = etree.SubElement(maps_root, 'Map')

    for attribute, val in map_attributes.items():
        map.set(attribute, val)
    return (maps_root, map)

def coord_to_str(coord: float, int_length: int) -> str:
    d = Decimal(str(coord))
    integer_part = int(d) - 360 if int(d) > 180 else int(d)
    integer_part_str = str(integer_part).zfill(int_length)
    fractional_part = d % 1
    fractional_part_str = f'{fractional_part:.3f}'.lstrip('0').lstrip('-0')
    leader = '+' if integer_part > 0 else ''
    return leader + integer_part_str + fractional_part_str

def long_to_str(coord: float) -> str:
    return coord_to_str(coord, 3)

def lat_to_str(coord: float) -> str:
    return coord_to_str(coord, 2)

def make_infill_xml(sigmet_poly: list[list[float]]) -> etree.Element:
    # Create the empty elements
    infill_element = etree.Element('Infill')  
    point_element = etree.SubElement(infill_element, 'Point')

    # Create all the point strings from the poly coords and add inside <Point> element
    point_strings = [lat_to_str(latitude) + long_to_str(longitude) for longitude, latitude in sigmet_poly]
    log('created polygon with ISO 6709 coordinates %s' % point_strings)
    point_element.text = '/'.join(point_strings)

    return infill_element

def run(vatsys_maps_dir: str, output_filename: str):
    
    log('running with output location %s' % Path(vatsys_maps_dir, output_filename))

    # Fetch SIGMETs from API
    try:
        r = requests.get(ISIGMET_API_URL)
        sigmets_json = r.json()
        log('fetched JSON from %s' % ISIGMET_API_URL)
    except Exception as e:
        error('could not fetch SIGMETs from API')
        traceback.print_exc()
        exit_with_wait()

    # Make the XML
    try:
        # Make the base <Maps> and <Map> element
        maps_root, map_element = make_base_map_xml()
        # Iterate over each KZAK GeoJSON feature and make the <Infill> xml. Add to <Map>
        filtered = filter_kzak_sigmets(sigmets_json)
        log('found %d SIGMETs for KZAK' % len(filtered))
        for geojson_sigmet in filtered:
            for poly_coords in geojson_sigmet['geometry']['coordinates']:
                infill_xml = make_infill_xml(poly_coords)
                map_element.append(infill_xml)
    except Exception:
        error('could not form XML')
        traceback.print_exc()
        exit_with_wait()

    
    # Write output XML
    try:
        path = Path(vatsys_maps_dir, output_filename)
        etree.ElementTree(maps_root).write(path, pretty_print=True)
        log('wrote XML file to %s' % path)
    except:
        error('could not write output file to %s' % path)
        traceback.print_exc()
        exit_with_wait()

if __name__ == '__main__':

    ## Creating the argument parser
    ## TODO: add options for verbosity? or to launch vatSys after?
    parser = argparse.ArgumentParser()
    parser.add_argument('--mapsdir', help="location of vatSys Maps folder for ATOP Oakland profile")
    parser.add_argument('--filename', help="full name of output XML file (including .xml)")
    parser.add_argument('--exec', help="location of vatSys executable")
    parser.add_argument('--color', help="name of vatSys color (from Colours.xml) to use for SIGMETs")
    args = parser.parse_args()

    # Get profile maps dir from command line first, or do auto. Fail out if we can't find
    maps_dir = args.mapsdir if args.mapsdir is not None else find_vatsys_maps_dir()
    if maps_dir is None:
        error('could not find suitable vatSys Maps folder for ATOP Oakland profile')
        exit_with_wait()
    
    # Get output filename for command line first, or just default
    filename = args.filename if args.filename is not None else DEFAULT_FILENAME

    # We've got the maps_dir and filename now, so we can run
    run(maps_dir, filename)
    
    # Get the vatSys executable to run after
    exec_path = args.exec if args.exec is not None else find_vatsys_exec()
    if exec_path is None:
        error('could not find suitable vatSys executable')
        exit_with_wait()
    else:
        log('opening vatSys executable at %s' % exec_path)
        subprocess.Popen([exec_path])
        exit_with_wait()