import xml.etree.ElementTree as ET
import xml.dom.minidom
import openpyxl
import os

def load_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    return root

def load_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    return sheet

def map_fields(xml_root, excel_sheet):
    field_mapping = {}
    print("Pola w pliku XML:")
    for child in xml_root[0]:
        print(child.tag)

    print("\nKolumny w pliku Excela:")
    for cell in excel_sheet[1]:
        print(cell.value)

    while True:
        xml_field = input("\nWybierz pole z pliku XML (lub 'koniec' aby zakończyć): ")
        if xml_field == "koniec":
            break
        excel_column = input("Wybierz kolumnę z pliku Excela: ")
        field_mapping[xml_field] = excel_column

    return field_mapping

def generate_mapped_xml(xml_root, field_mapping, excel_sheet):
    root = ET.Element(xml_root.tag)

    column_indices = {cell.value: cell.column for cell in excel_sheet[1]}  # Mapowanie nazw kolumn Excela na ich indeksy

    for row in excel_sheet.iter_rows(min_row=2, values_only=True):
        element = ET.SubElement(root, xml_root[0].tag)
        for xml_field, excel_column in field_mapping.items():
            try:
                if excel_column in column_indices:
                    column_index = column_indices[excel_column]
                    cell_value = row[column_index - 1]
                    sub_element = ET.SubElement(element, xml_field)
                    sub_element.text = str(cell_value)
                else:
                    print(f"Kolumna '{excel_column}' nie istnieje w arkuszu Excela.")
            except ValueError:
                print(f"Błąd podczas przetwarzania kolumny '{excel_column}'.")
                continue

    mapped_xml_file = "mapped_output.xml"
    if os.path.exists(mapped_xml_file):
        os.remove(mapped_xml_file)

    xml_string = ET.tostring(root, encoding='utf-8')
    parsed_xml = xml.dom.minidom.parseString(xml_string)
    pretty_xml = parsed_xml.toprettyxml(indent="  ")
    with open(mapped_xml_file, "w") as file:
        file.write(pretty_xml)


# Wczytanie plików XML i Excela
xml_file_path = "XML From Excel - Map.xml"
excel_file_path = "XML From Excel - Not Mapped.xlsx"

xml_root = load_xml(xml_file_path)
excel_sheet = load_excel(excel_file_path)

# Mapowanie pól
field_mapping = map_fields(xml_root, excel_sheet)

# Generowanie pliku XML z mapowaniem
generate_mapped_xml(xml_root, field_mapping, excel_sheet)
