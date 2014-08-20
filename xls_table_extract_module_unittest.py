""" Unittest for xls_table_extract_module.


"""

import os, time, re, sys
import unittest
from xls_table_extract_module import XlsExtractor

class XlsExtractorTestCase(unittest.TestCase):
    """ Base class for all XlsExtractor test

    """
    def setup_basic_parameters(self):
        """ Set up essential parameters for the main class.

        """

class XlsExtractorTest(XlsExtractorTestCase):
    """ Main unit test for the module
        Split to two separate test object??

    """

    def setUp(self):
        """
            Create mulitple instances of different xlsExtactor
            
        """
        self.xls_set_class_1 = XlsExtractor(fname = r'C:\Python27\Lib\site-packages\excel_table_extract\testset.xls', sheetname= 'Sheet1',
                             param_start_key = 'start//', param_end_key = 'end//',
                             header_key = '', col_len = 3)

        self.xls_set_class_1.initialize_excel()

    def test_calculate_start_end_pos_for_data(self):

        start_address_row, start_address_col, end_address_row = self.xls_set_class_1.calculate_start_end_pos_for_data()
        self.assertTupleEqual((start_address_row, start_address_col, end_address_row), (7,4,11))

    def test_get_raw_block_fr_xls(self):
        start_address_row, start_address_col, end_address_row = self.xls_set_class_1.calculate_start_end_pos_for_data()
        s = self.xls_set_class_1.get_raw_block_fr_xls(start_address_row, start_address_col, end_address_row, self.xls_set_class_1.col_len)
        self.assertListEqual(s, [[None, 'label1', 2.0, 3.0], ['#', 'label2', 5.0, 6.0], [None, None, None, None], [None, 'label3', 8.0, 9.0]])   

    def test_filter_space_and_comment_row(self):
        s = self.xls_set_class_1.filter_space_and_comment_row([[None, 1.0, 2.0, 3.0],['#', 4.0, 5.0, 6.0], [None, None, None, None], [None, 7.0, 8.0, 9.0]])
        self.assertListEqual(s, [[None, 1.0, 2.0, 3.0], [None, 7.0, 8.0, 9.0]])

    def test_exclude_comment_block(self):
        s = self.xls_set_class_1.exclude_comment_block([[None, 1.0, 2.0, 3.0], [None, 7.0, 8.0, 9.0]])
        self.assertListEqual(s, [[1.0, 2.0, 3.0], [7.0, 8.0, 9.0]])
        
    def tearDown(self):
        """ For each instances try to close it.
        """
        try:
            self.xls_set_class_1.close_excel()
        finally:
            pass
    
if __name__ == '__main__':

    unittest.main()
