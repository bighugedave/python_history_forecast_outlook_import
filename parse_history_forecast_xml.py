import xml.etree.ElementTree as ET
import os


class ParseHistoryForecastXML:
    def __init__(self, xml_dir):
        self.xml_dir = xml_dir
