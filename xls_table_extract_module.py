'''
#############################################

 xls table extract module
 Author: Tan Kok Hua (Guohua tan)
 Email: spider123@gmail.com
 Revised date: Aug 2014
 
##############################################

Usage:
    Will replace the xls setting extract module.
    Underconstruction

Updates:
    Aug 11 2014: Finish on the header tag.
    Aug 10 2014: Modified from xls_setting_extract_module.

TODO:
    make the columna and row position as variable 
    give error if cannot find the tag
    To review whether want to put table number in tag --> as although table got two, user can have preference to select 


'''
import sys
import os
import re
try:
    from pyExcel import UseExcel
except:
    ## 'Use alternative import if pyExcel is not present'
    from pyET_tools.pyExcel import UseExcel

class XlsExtractor(object):
    def __init__(self, fname = '', sheetname= '', param_start_key = '', param_end_key = '', header_key = '', col_len = 3):
        """ Extract setting info from xls.
        """
        ## Parameters for get_col_list_fr_xls (inputs)
        self.keyword_fname              = fname
        self.sheetname                  = sheetname
        self.param_start_key_tag        = param_start_key
        self.param_end_key_tag          = param_end_key
        self.header_key_tag             = header_key
        self.col_len                    = col_len

        ## Excel object
        self.excel_obj                  = object()
        
        ## Outputs-- different type of outputs for further processing.
        ## current set for one table
        self.full_filtered_data         = list()
        self.header_list                = list()
        self.data_label_list            = list()
        self.data_value_list            = list()
        self.label_value_dict           = dict()

    def set_file_and_sheet_name(self, xls_plus_shtname_list):
        '''Set the xls file and sheet name'''
        self.keyword_fname = xls_plus_shtname_list[0]
        self.sheetname = xls_plus_shtname_list[1]

    def set_start_and_end_tag(self, xls_tag_border_list):
        self.param_start_key_tag = xls_tag_border_list[0]
        self.param_end_key_tag = xls_tag_border_list[1]

    def process_start_tag(self, tag_label):
        """ Process the start tag for addtional attributes. Particularly after the #
            The # is used in header tag to give indication on the number of columns.
            Args:
                tag_label (str): start tag.
            Returns:
                (int): attributes number. If do not have #, return default value 1.
        """
        tag_strlist = tag_label.split('#')
        if len(tag_strlist) == 2:
            return int(re.match('.*#(\d+)',tag_label).group(1))
        else:
            return 1


    def initialize_excel(self):
        """ Initialize the excel setting.
            Required a valid filename.
        """
        self.excel_obj = UseExcel(self.keyword_fname)

    def close_excel(self):
        """ Close the excel object.
        """
        self.excel_obj.close()

    def open_excel_and_process_block_data(self):
        """ Consolidated the function necessary to open excel and get raw block.
            Take care of the excel function in this section.
            Will process header if header_tag is present.
            Returns:
                (list): list of data with the space row and comment being filtered out.

            Note: currently only for one table
        """
        self.initialize_excel()
        try:
            
            ## processed header block if present
            if not self.header_key_tag == '':
                no_of_header_cols = self.process_start_tag(self.header_key_tag)
                header_start_add_row, header_start_add_col = self.get_position_of_tag_in_xls(self.header_key_tag)
                self.header_list = self.get_raw_block_fr_xls(header_start_add_row-1, header_start_add_col +1, header_start_add_row, no_of_header_cols-1)[0]

            ## processed data block
            start_address_row, start_address_col, end_address_row = self.calculate_start_end_pos_for_data()
            s = self.get_raw_block_fr_xls(start_address_row, start_address_col, end_address_row, self.col_len)
        except:
            print 'Problem getting data from excel'
            raise
        finally:
            self.close_excel()

        s = self.filter_space_and_comment_row(s)
        self.full_filtered_data  = self.exclude_comment_block(s)
        self.data_label_list, self.data_value_list  = self.segregate_to_label_values_block(self.full_filtered_data )

        self.create_label_value_dict()

    def calculate_start_end_pos_for_data(self):
        """ Locate all the table position. Get start row, column and end row.
            Find the tag position. Need adjust for the data location.
            Excel object need to be present.
            Adjusted for default header
            
            <tag start>
            header
            comment col, data
            comment col, data2
            ...
            <tag end>

            Returns:
                start_address_row (int): start row (tag position + 1)
                start_address_col (int): 
                end_address_row (int): end row (tag position -1)
            
        """
        start_address_row = self.excel_obj.find_keyword(self.sheetname, self.param_start_key_tag)[1][0] + 1
        start_address_col = self.excel_obj.find_keyword(self.sheetname, self.param_start_key_tag)[1][1]
        end_address_row = self.excel_obj.find_keyword(self.sheetname, self.param_end_key_tag)[1][0] - 1

        ## end row must be greater than start row to ensure have data.
        assert end_address_row > start_address_row

        return start_address_row, start_address_col, end_address_row

    def get_position_of_tag_in_xls(self, tag_label):
        """ Get the row and column position of the tag in the excel.
            Excel obj must exist.
            Args:
                tag_label (str): target tag to processed.
            Returns:
                (int), (int):  row and column. (row, column)
        """
        start_address_row = self.excel_obj.find_keyword(self.sheetname, tag_label)[1][0]
        start_address_col = self.excel_obj.find_keyword(self.sheetname, tag_label)[1][1]
        return start_address_row, start_address_col
        

    def get_raw_block_fr_xls(self, start_address_row, start_address_col, end_address_row, num_of_col):
        """ Get one single block fr excel bound by the start and end tag.
            This is the raw block without any filtering.
            It can be used to get comment block or any block within the pos stated.
            Excel object need to be present.

            Args:
                start_address_row (int): start row (tag position + 1)
                start_address_col (int): start col.
                end_address_row (int): end row (tag position -1)
                num_of_col (int): num of column to grab from the start col. <--make it as end of column??

        """
        dataset = self.excel_obj.getrange(self.sheetname, (start_address_row+1,start_address_col, end_address_row, start_address_col +num_of_col))
        dataset = [list(n) for n in dataset]
        return dataset

    def filter_space_and_comment_row(self, full_raw_data_block):
        """ Filter the raw_data_block for those that is designated as space or being commmented.
            The full data include the comment column.
            Two criteria here:
                First:
                    Treat as space when the first data column is treated as empty.
                    Take in full raw data, check if first data column n[1] (or label column) excluding the comment block is empty.
                    If yes, treat it as space block.
                2nd:
                    Treat as comment when the comment block or the n[0] column are marked with '#'.
            Args:
                full_raw_data_block (list): list of rows of row data.
                                            The full_raw_data_block include the comment block plus data

            Returns:
                (list): list of data with the space row and comment being filtered out.
        """
        return [n for n in full_raw_data_block if n[1] is not None and n[0] != '#']

    def exclude_comment_block(self, filtered_full_data_block):
        """ Filter out the comment block after the data being processed.
            This method is call after the filter_space_and_comment_row function is being execute.
            It remove the comment block from the data block.
            Args:
                filtered_full_data_block (list): post processed data block for comment and space.
            Returns:
                filtered_data_block (list): return the remaining portion without the data block.
        """
        return [ n[1:] for n in filtered_full_data_block ]

    def segregate_to_label_values_block(self, filtered_data_block):
        """
            Segregate the filtered_data_block into label (first data column) and values (Subsequent data column).
            The args passed in must be already filtered and stripped of the comment block.
            The function should come after the exclude_comment_block function.
            Args:
                filtered_data_block (list): fitered data block.
            Returns:
                label (list): single list. First column
                values (list): can be single list or list of list. Subsequent columns.
        """
        return [n[0] for n in filtered_data_block], [n[1:] for n in filtered_data_block]

    def create_label_value_dict(self):
        """ Create label value pair after segregating the filtered data block. Foramt in dict.
            self.data_label_list and self.data_value_list must not be empty.
            Write to self.label_value_dict.
        """
        for label, data in zip(self.data_label_list, self.data_value_list):
            self.label_value_dict[label] =  data
            
    def get_header_fr_xls(self):
        """ Get header label for the particular table. Retrieved using the header tag.
            Note the header tag is written with # indicating the number of header columns.
            Header data is at the same position as the tag and occupy only one row.
            <Header tag> Header 1, Header 2.....
        """
        no_of_header_cols = self.process_start_tag(self.header_key_tag)
        
        self.initialize_excel()
        try:
            start_address_row, start_address_col = self.get_position_of_tag_in_xls(self.header_key_tag)
            ## start address neeed -1 because the function will auto increase the start address by 1
            ## end address col -1 because the first column is data included.
            self.header_list = self.get_raw_block_fr_xls(start_address_row-1, start_address_col +1, start_address_row, no_of_header_cols-1)[0]
            self.header_list = [n.encode() for n in self.header_list ]
            return self.header_list
        except:
            print 'Problem getting header from excel'
            raise
        finally:
            self.close_excel()


if __name__ == '__main__':
    test = [9,11]

    if 9 in test:
        """"""
        xls_set_class = XlsExtractor(fname = r'C:\Python27\Lib\site-packages\excel_table_extract\testset.xls', sheetname= 'Sheet1',
                                     param_start_key = 'start//', param_end_key = 'end//',
                                     header_key = 'header#3//', col_len = 3) 
        
    if 10 in test:
        """

        """
        xls_set_class.initialize_excel()
        try:
            start_address_row, start_address_col, end_address_row = xls_set_class.calculate_start_end_pos_for_data()
            print start_address_row, start_address_col, end_address_row
            s = xls_set_class.get_raw_block_fr_xls(start_address_row, start_address_col, end_address_row, xls_set_class.col_len)
            print s
            s = xls_set_class.filter_space_and_comment_row(s)
            print s
            s = xls_set_class.exclude_comment_block(s)
            print s
            s = xls_set_class.segregate_to_label_values_block(s)
            print s

        finally:
            print
            xls_set_class.close_excel()

    if 11 in test:
        xls_set_class.open_excel_and_process_block_data()
        
        print xls_set_class.data_label_list
        ## >>> [u'label1', u'label3']
        
        print xls_set_class.data_value_list
        ## >>> [[2.0, 3.0], [8.0, 9.0]]
        
        print xls_set_class.label_value_dict
        ## >>> {u'label1': [2.0, 3.0], u'label3': [8.0, 9.0]}

        print xls_set_class.header_list
        ## >>> [u'header1', u'header2', u'header3']




