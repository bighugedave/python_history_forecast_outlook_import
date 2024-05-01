import xml.etree.ElementTree as ET
import os


class ParseHistoryForecastXML:
    def __init__(self, xml_dir):
        self.ext = '.xml'
        self.xml_dir = xml_dir
        self.xml_files = []

    def get_xml_files(self):
        for xml_file in os.listdir(self.xml_dir):
            if xml_file.lower().endswith(self.ext):
                print(xml_file)
                self.xml_files.append(xml_file)
            else:
                continue
        self.parse_xml_files()

    def parse_xml_files(self):
        for xml_file in self.xml_files:
            tree = ET.parse('{}\\{}'.format(self.xml_dir, xml_file))
            root = tree.getroot()
            for child in root.findall(".//G_CONSIDERED_DATE"):
                print(child.find("CHAR_CONSIDERED_DATE").text)
