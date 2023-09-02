# Find head data using serial number
# it's necessary need right serial number
from HeadData import get_head_data


# import all local files parsers
from SPTParser import SPTParser
from TMKParser import TMKParser


class ReportCreator:
    """ 
    This is report builder, that use 
    report and head data from lists 
    and build right xsls file
    """

    def __init__(self):
        pass

    def build(self, head_data, parser_list):
        pass
