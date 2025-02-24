import pandas as pd
import os
from typing import Dict
from datetime import datetime
from FolderManager import FolderManager

class ReportFurnishing:

    def __init__(self, 
                 folder_manager: FolderManager):
        self.folder_manager = folder_manager
        self.manufacturer = folder_manager.manufacturer
        self.contract = folder_manager.contract
        self.output_file_path = folder_manager.get_folder_path('output')
        self.datesig = self.datesig = datetime.today().strftime('%Y%m%d')
        self.color_pane = {'light_blue': '#D7ECFF',
                           'yellow': '#FFD500',
                           'grey': '#808080',
                           'pink': '#FFC0CB',
                           'white': '#FFFFFF'}
        self.header_format_ccx = {'bold': True,
                                  'text_wrap': True,
                                  'valign': 'vcenter',
                                  'align': 'center',
                                  'fg_color': self.color_pane['yellow'],
                                  'border': 1}
        self.header_format_tp = {'bold': True,
                                 'text_wrap': True,
                                 'valign': 'vcenter',
                                 'align': 'center',
                                 'fg_color': self.color_pane['grey'],
                                 'font_color': self.color_pane['white'],
                                 'border': 1}
        self.header_format_custom = {'bold': True,
                                     'text_wrap': True,
                                     'valign': 'vcenter',
                                     'align': 'center',
                                     'fg_color': self.color_pane['light_blue'],
                                     'border': 1}
        self.header_format_clear = {'bold': True,
                                    'text_wrap': True,
                                    'valign': 'vcenter',
                                    'align': 'center',
                                    'border': 1}
        self.cell_format_warning = {'bg_color': self.color_pane['pink']}

        self.report_header_dict = {'dedeup': ['Mfg Part Num', 
                                              'Vendor Part Num',
                                              'Buyer Part Num',
                                              'Description',
                                              'Contract Price',
                                              'UOM',
                                              'QOE',
                                              'Effective Date',
                                              'Expiration Date',
                                              'Contract Number (Old)',
                                              'Contract Line',
                                              'Source System',
                                              'Action',
                                              'Mfg Part Num (New)',
                                              'Vendor Part Num (New)',
                                              'Description (New)',
                                              'Unit Price (New)',
                                              'UOM (New)',
                                              'QOE (New)',
                                              'Contract Number (New)',
                                              'Same UOM',
                                              'Same QOE',
                                              'EA Cost Diff',
                                              'Desc. Similarity'],
                                   'itemmast': ['Contract Number',
                                                'Mfg Part Num',
                                                'Vendor Part Num',
                                                'Description',
                                                'Contract Price',
                                                'UOM',
                                                'QOE',
                                                'seq',
                                                'Effective Date',
                                                'Expiration Date',
                                                'Mfg Part Num (Reduced)',
                                                'Description (Matched Item)',
                                                'Desc. Similarity',
                                                'Mfg Part Num (Infor)',
                                                'Item',
                                                'Item Type',
                                                'UOM Conversion (Infor)',
                                                'Valid For Buying',
                                                'All Valid BuyUOM and CF',
                                                'Item UOM Check Result',
                                                'Numbers of Item Matched'],
                                   'replace': ['Contract Number',
                                               'Mfg Part Num',
                                               'Vendor Part Num',
                                               'Description',
                                               'Contract Price',
                                               'UOM',
                                               'QOE',
                                               'Effective Date',
                                               'Expiration Date',
                                               'seq',
                                               'On Replacement Contract',
                                               'Item Type'],
                                    'dedup_summary': ['Source System',
                                                      'Contract Number',
                                                      'Manufacturer Name (CCX)',
                                                      'Total Line Count',
                                                      'Overlapping Line Count']
                                                      }

    def make_dedup_report(self, 
                          df_dedup: Dict[str, pd.DataFrame],
                          df_summary: pd.DataFrame,
                          df_raw: pd.DataFrame):
        
        file_name = f"dedup_output_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx"
        
        df_summary.columns = self.report_header_dict['dedup_summary']   
        with pd.ExcelWriter(os.path.join(self.output_file_path, file_name), engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Summary']

            header_format_ccx = workbook.add_format(self.header_format_ccx)
            header_format_tp = workbook.add_format(self.header_format_tp)
            header_format_custom = workbook.add_format(self.header_format_custom)
            header_format_clear = workbook.add_format(self.header_format_clear)

            header_styles_summary = {}
            for col in range(0, 5):
                header_styles_summary[col] = header_format_clear

            for col_num, fmt in header_styles_summary.items():
                worksheet.write(0, col_num, df_summary.columns[col_num], fmt)
            
            worksheet.autofilter(0, 0, df_summary.shape[0], df_summary.shape[1]-1)

            for key, df in df_dedup.items():
                df.columns = self.report_header_dict['dedeup']
                df.to_excel(writer, sheet_name=key, index=False)
                worksheet = writer.sheets[key]
                header_styles = {}
                for col in range(0, 11):
                    header_styles[col] = header_format_ccx
                for col in range(11, 13):
                    header_styles[col] = header_format_clear
                for col in range(13, 20):
                    header_styles[col] = header_format_tp
                for col in range(20, 24):
                    header_styles[col] = header_format_custom

                for col_num, fmt in header_styles.items():
                    worksheet.write(0, col_num, df.columns[col_num], fmt)
            
                cell_format_warning = workbook.add_format(self.cell_format_warning)
                worksheet.conditional_format(1, 0, df.shape[0], df.shape[1]-1, {'type': 'cell',
                                                                          'criteria': '==',
                                                                          'value': '"Review"',
                                                                          'format': cell_format_warning})
                worksheet.autofilter(0, 0, df.shape[0], df.shape[1])
            
            df_raw.to_excel(writer, sheet_name='Raw', index=False)
        
        return "Duplication report generated."
    
    
    def make_itemmast_report(self, 
                             df: pd.DataFrame,
                             sheet_name: str = "IM Match"):
        file_name = f"itemmast_match_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx"
        df.columns = self.report_header_dict['itemmast']
        with pd.ExcelWriter(os.path.join(self.output_file_path, file_name), engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            header_format_tp = workbook.add_format(self.header_format_tp)
            header_format_custom = workbook.add_format(self.header_format_custom)
            header_format_clear = workbook.add_format(self.header_format_clear)

            header_styles = {}
            for col in range(0, 10):
                header_styles[col] = header_format_tp
            for col in range(10, 14):
                header_styles[col] = header_format_clear
            for col in range(14, 21):
                header_styles[col] = header_format_custom

            for col_num, fmt in header_styles.items():
                worksheet.write(0, col_num, df.columns[col_num], fmt)
            
            cell_format_warning = workbook.add_format(self.cell_format_warning)
            worksheet.conditional_format(1, 0, df.shape[0], df.shape[1], {'type': 'cell',
                                                                          'criteria': '==',
                                                                          'value': '"Failed"',
                                                                          'format': cell_format_warning})
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
        
        return "Itemmast report generated."
    
    def make_replace_report(self,
                            df: pd.DataFrame,
                            sheet_name: str = "NoReplacement"):
        file_name = f"replacement_leftover_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx"
        df.columns = self.report_header_dict['replace']
        with pd.ExcelWriter(os.path.join(self.output_file_path, file_name), engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            header_format_ccx = workbook.add_format(self.header_format_ccx)
            header_format_clear = workbook.add_format(self.header_format_clear)

            header_styles = {}
            for col in range(0, 10):
                header_styles[col] = header_format_ccx
            for col in range(10, 13):
                header_styles[col] = header_format_clear

            for col_num, fmt in header_styles.items():
                worksheet.write(0, col_num, df.columns[col_num], fmt)
            
            cell_format_warning = workbook.add_format(self.cell_format_warning)
            worksheet.conditional_format(1, 0, df.shape[0], df.shape[1], {'type': 'cell',
                                                                          'criteria': '==',
                                                                          'value': '"Immast"',
                                                                          'format': cell_format_warning})
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
        return "Replacement leftover report generated."