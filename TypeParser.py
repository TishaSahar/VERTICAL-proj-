import pyexcel

from HeadData import get_head_data


class TypeParser:
    """
    This is the type parser
    that use list of all files
    and chose type of file using head 
    data parser. 
    """

    # this touple contains the result of type parsing
    data_touple = {'ВЗЛЕТ': [], 'ВКТ': [], 'МКТС': [], 'СПТ': [], 'ТВ-7': [], 'ТМК': []}

    def __init__(self, data_list):
        for file in data_list:
            # try to get file with .xlsx extention
            file = self.file_type_checker(file)
            head_data = get_head_data(file)

    
    def build(self, all_parsing_files):
        pass


    def file_type_checker(self, file):
        if '.xls' in file:         
            if '.xslx' not in file:
                file += 'x'
                pyexcel.save_book_as(file_name=file,
                                dest_file_name=file)
            return file
        else:
            raise "Input type error"
        