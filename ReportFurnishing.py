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
                   'white': '#FFFFFF',
                   'green': '#C6E0B4',
                   'gold': '#DAA520'}
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
                                               'Buyer Part Num',
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
                worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
            
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
            worksheet.conditional_format(1, 0, df.shape[0], df.shape[1]-1, {'type': 'cell',
                                                                          'criteria': '==',
                                                                          'value': '"Failed"',
                                                                          'format': cell_format_warning})
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
        
        return "Itemmast report generated."
    
    def make_replacement_report(self,
                            df: pd.DataFrame,
                            df2: pd.DataFrame,
                            sheet_name: str = "NoReplacement",
                            sheet_name2: str = "Comparison"):
        file_name = f"replacement_comparison_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx"
        df.columns = self.report_header_dict['replace']
        comparison_columns = ['MFN',
                              'Contract Number_TP',
                              'VN_TP',
                              'IN_TP',
                              'Description_TP',
                              'UnitCost_TP',
                              'UOM_TP',
                              'QOE_TP',
                              'Effective Date_TP',
                              'Expiration Date_TP',
                              'FileName_TP',
                              'seq_TP',
                              'Contract Number_CCX',
                              'VN_CCX',
                              'IN_CCX',
                              'Description_CCX',
                              'UnitCost_CCX',
                              'UOM_CCX',
                              'QOE_CCX',
                              'Effective Date_CCX',
                              'Expiration Date_CCX',
                              'FileName_CCX',
                              'seq_CCX',
                              '_merge']
        missing_columns = [col for col in comparison_columns if col not in df2.columns]
        if missing_columns:
            raise KeyError(f"Missing expected comparison columns: {missing_columns}")
        rename_map = {'MFN': 'Manufacturer Part Num',
                      'Contract Number_TP': 'Contract Number (TP)',
                      'VN_TP': 'Vendor Part Num (TP)',
                      'IN_TP': 'Buyer Part Num (TP)',
                      'Description_TP': 'Description (TP)',
                      'UnitCost_TP': 'Unit Cost (TP)',
                      'UOM_TP': 'UOM (TP)',
                      'QOE_TP': 'QOE (TP)',
                      'Effective Date_TP': 'Effective Date (TP)',
                      'Expiration Date_TP': 'Expiration Date (TP)',
                      'FileName_TP': 'FileName (TP)',
                      'seq_TP': 'seq (TP)',
                      'Contract Number_CCX': 'Contract Number (CCX)',
                      'VN_CCX': 'Vendor Part Num (CCX)',
                      'IN_CCX': 'Buyer Part Num (CCX)',
                      'Description_CCX': 'Description (CCX)',
                      'UnitCost_CCX': 'Unit Cost (CCX)',
                      'UOM_CCX': 'UOM (CCX)',
                      'QOE_CCX': 'QOE (CCX)',
                      'Effective Date_CCX': 'Effective Date (CCX)',
                      'Expiration Date_CCX': 'Expiration Date (CCX)',
                      'FileName_CCX': 'FileName (CCX)',
                      'seq_CCX': 'seq (CCX)'}
        df2_formatted = df2[comparison_columns].copy()
        df2_formatted.rename(columns=rename_map, inplace=True)
        
        with pd.ExcelWriter(os.path.join(self.output_file_path, file_name), engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            header_format_ccx = workbook.add_format(self.header_format_ccx)
            header_format_clear = workbook.add_format(self.header_format_clear)
            header_format_green = workbook.add_format({**self.header_format_clear,
                                                       'fg_color': self.color_pane['green']})
            header_format_gold = workbook.add_format({**self.header_format_clear,
                                                      'fg_color': self.color_pane['gold']})

            header_styles = {}
            for col in range(0, 10):
                header_styles[col] = header_format_ccx
            for col in range(10, 13):
                header_styles[col] = header_format_clear

            for col_num, fmt in header_styles.items():
                worksheet.write(0, col_num, df.columns[col_num], fmt)
            
            cell_format_warning = workbook.add_format(self.cell_format_warning)
            worksheet.conditional_format(1, 0, df.shape[0], df.shape[1]-1, {'type': 'cell',
                                                                          'criteria': '==',
                                                                          'value': '"Immast"',
                                                                          'format': cell_format_warning})
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)

            df2_formatted.to_excel(writer, sheet_name=sheet_name2, index=False)
            worksheet_comparison = writer.sheets[sheet_name2]

            for col_num, column_name in enumerate(df2_formatted.columns):
                if column_name.endswith('(TP)'):
                    worksheet_comparison.write(0, col_num, column_name, header_format_green)
                elif column_name.endswith('(CCX)'):
                    worksheet_comparison.write(0, col_num, column_name, header_format_gold)
                else:
                    worksheet_comparison.write(0, col_num, column_name, header_format_clear)

            worksheet_comparison.autofilter(0, 0, df2_formatted.shape[0], df2_formatted.shape[1]-1)
        return "Replacement leftover report generated."