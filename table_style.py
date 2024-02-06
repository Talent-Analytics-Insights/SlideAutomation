"""Tablestyle documentation:
    https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.TableStyle
    https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablestyle?view=openxml-2.8.1
    
"""
import os
import shutil
import zipfile
import uuid  

from . import style_element

  
def generate_guid() -> str:  
    """Function to generate a random GUID (Globally Unique Identifier) in the proper format. See microsoft docs for this format.
       https://learn.microsoft.com/en-us/dynamicsax-2012/developer/guids

    Returns:
        str: Valid GUID
    """
    # Generate a random UUID  
    guid = uuid.uuid4()  
  
    # Convert the UUID to a string  
    guid_str = str(guid)  
  
    # Format the GUID according to the Microsoft Office guidelines  
    guid_str = "{" + guid_str.upper() + "}"  
  
    return guid_str

def create_new_table_style(style_id: str, style_name: str, style_header_color: str, style_row_even_color: str, style_row_odd_color: str, include_first_row: bool = True) -> str:
    """Function to create new style element for table

    Args:
        style_id (str): GUID to give to new style element
        style_name (str): Style Name (this will appear in PowerPoint when selecting the Table Style)
        style_header_color (str): Color Code of header
        style_row_even_color (str): Color Code of even rows
        style_row_odd_color (str): Color Code of odd rows
        include_first_row (bool, optional): Whether to include a first row. Defaults to True.

    Returns:
        str: XML string containing Style Element
    """

    # Create and return element
    style = style_element.get_style_element(
        style_id, style_name, style_header_color, style_row_even_color, style_row_odd_color, include_first_row)
    return style


def add_style_to_powerpoint(ppt_file_in: str, ppt_file_out: str, style_name: str, style_header_color: str, style_row_even_color: str, style_row_odd_color: str, include_first_row: bool = True):
    """Function to add a valid Table Style XML to a PowerPoint file. A template file should be provided

    Args:
        ppt_file_in (str): PowerPoint file to add the style to. Can be an empty presentation
        ppt_file_out (str): File to store PowerPoint. If the same as ppt_file_in, this will be overwritten
        style_name (str): Style Name (this will appear in PowerPoint when selecting the Table Style)
        style_header_color (str): Color Code of header
        style_row_even_color (str): Color Code of even rows
        style_row_odd_color (str): Color Code of odd rows
        include_first_row (bool, optional): Whether to include a first row. Defaults to True.

    Returns:
        str: The GUID of the created table style
    """
    # Create tmp folder, required to extract and alter the ppt
    if not os.path.exists('tmp'):
        os.mkdir('tmp')
    
    # Make temp zip file
    if not os.path.exists('tmp/ppt_zip_temp'):
        os.mkdir('tmp/ppt_zip_temp')

    style_id = generate_guid()

    # Open and extract input ppt file as zip file
    with zipfile.ZipFile(ppt_file_in, 'r') as zip_ref:
        zip_ref.extractall('tmp/ppt_zip_temp')

    # Path to the tableStyles file in the unzipped ppt
    table_styles_path = "tmp/ppt_zip_temp/ppt/tableStyles.xml"

    # Create the new style
    new_table_style = create_new_table_style(
        style_id, style_name, style_header_color, style_row_even_color, style_row_odd_color, include_first_row)

    # Read the current style
    with open(table_styles_path, 'r') as f:
        current_table_styles = f.read()
    
    # If there is no style, create a tblStyleLst and embed the new style
    if len(current_table_styles) < 250:
        new_table_styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{style_id}">{new_table_style}</a:tblStyleLst>'
    # If there were already defined styles, add new style
    else:
        new_table_styles = "<".join(current_table_styles.split('<')[:3]) + new_table_style + "<" + "<".join(current_table_styles.split('<')[3:])

    # Write the tableStyles.xml file
    with open(table_styles_path, 'w') as f:
        f.write(new_table_styles)

    # create a new PowerPoint file with the modified contents
    with zipfile.ZipFile(ppt_file_out, 'w') as zip_ref:
        for root, dirs, files in os.walk('tmp/ppt_zip_temp'):
            for file in files:
                file_path = os.path.join(root, file)
                zip_ref.write(file_path, os.path.relpath(file_path, 'tmp/ppt_zip_temp'))
    
    # Remove tmp folder
    shutil.rmtree('tmp')
    
    return style_id
